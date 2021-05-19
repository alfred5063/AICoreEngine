
from transformation.mergeQueries import merge
from transformation.df_preprocessing import preprocessing
from transformation.manageRow import remove_duplicate
from utils.audit_trail import audit_log
from utils.logging import logging
from transformation.ClassifyDuplicates import remove_excess_column

def check_matches_with_priority(raw_df, marc_result, base):
  try:
    audit_log('check_matches_with_priority', 'Custom checking logic for principal columns', base)
    principal_priority = [
      {'excel_column': 'Principal NRIC', 'marc_column': 'Principal IC'},
      {'excel_column': 'Principal Other Ic', 'marc_column': 'Principal Other IC'},
      {'excel_column': 'Principal Ext Ref Id (aka Client)', 'marc_column': 'Principal Ext Ref'},
      {'excel_column': 'Principal Name', 'marc_column': 'Principal Name MARC'},
      {'excel_column': 'Principal Int Ref Id (aka AAN)', 'marc_column': 'imppPrincipalId'}
    ]

    df = raw_df
    joined_result = raw_df
    for i, mapper in enumerate(principal_priority):
      #1. check null excel
      df = preprocessing(df, base, mapper['excel_column'])
      df[mapper['excel_column']] = df[mapper['excel_column']].astype(str)
      marc_result[mapper['marc_column']] = marc_result[mapper['marc_column']].astype(str)
      #2. merge with df
      joined_marc_result = merge(df, marc_result, base, how='left', left_on=mapper['excel_column'], right_on=mapper['marc_column'], indicator=True)
      joined_marc_result = remove_duplicate(joined_marc_result, 'No', keep='first')
      not_found = joined_marc_result[joined_marc_result['_merge'] == 'left_only']
      not_found = remove_excess_column(raw_df, not_found, base).drop(columns='_merge')
      joined_result = merge(joined_result, joined_marc_result[['No', '_merge']],base, how='left', on='No')
      joined_result['_merge'] = joined_result['_merge'] == 'both'
      #3. stamp column for matching
      joined_result.rename(columns={'_merge': f'_merge_{i}'}, inplace=True)
      #4. null_df are passed to the next round
      df = not_found # pass not_found to the next checking criteria
      #If we found the value for NRIC, then use NRIC to validate and 2-5 will be skip.
      #If NRIC not found, then we will use Other Ic to validate.
    #5. Combine result of all 5 columns
    joined_result['_merge'] = joined_result['_merge_0'] | joined_result['_merge_1'] | joined_result['_merge_2'] | joined_result['_merge_3'] | joined_result['_merge_4']
    return joined_result
  except Exception as error:
    logging('check_matches_with_priority', error, base)
    raise error
