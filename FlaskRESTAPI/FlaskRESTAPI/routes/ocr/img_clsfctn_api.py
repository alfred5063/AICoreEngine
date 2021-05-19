
from flask import current_app as app
from flask import request
from flaskr.response.response import Success, Failure
from flaskr.decorators.login_required import login_required
import os
from pathlib import Path
import sys

current_path = os.path.dirname(os.path.abspath(__file__))
ocr_path = str(Path(current_path).parent.parent.parent.parent) + '\\OCR.DTI'
sys.path.append(ocr_path)

from workflow.workflow_image_clssfctn import process_img_clssfctn


@app.route('/ocr_image_classification', methods = ['POST'])
@login_required
def process_image_clssfctn():
  try:
    img_path  = request.form['img_path']
    status, result = process_img_clssfctn(img_path)
    
    return Success({'message':status, 'result':str(result)})

  except Exception as error:
    return Failure({ 'message': str(error) }, debug = True )


