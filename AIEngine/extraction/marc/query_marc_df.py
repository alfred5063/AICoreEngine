from loading.query.query_as_df import query_as_df, sp_query_as_df

def query_marc_as_df(query, params, base):
  return query_as_df(query, params, MySqlConnector, base)

def sp_query_marc_as_df(query, params, base):
  return sp_query_as_df(query, params, MySqlConnector,base)
