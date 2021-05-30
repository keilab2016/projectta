from flask import Flask, request, send_file, render_template
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from workreport import getAllData, getAllID, findData, writeData
import os
import re

# If `entrypoint` is not defined in app.yaml, App Engine will look for an app
# called `app` in `main.py`.
app = Flask(__name__)


@app.route('/')
def hello():
    """Return a friendly HTTP greeting."""
    alldata = getAllData('102Fbm3HKFOeCg6Q0CEAQE88CLQ3xvtW3bTPDGydx02A',"2021前期")
    ids = getAllID(alldata)
    return render_template("/index.html", ids=ids, months=[5,6,7])

@app.route('/create')
def create():
    stuid = request.args['id']
    month = request.args['month']
    alldata = getAllData('102Fbm3HKFOeCg6Q0CEAQE88CLQ3xvtW3bTPDGydx02A',"2021前期")
    data = findData(stuid, int(month), alldata)
    file = writeData(stuid, month, data)
    if os.path.exists(file):
        XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        fname = re.sub(r'.*/','',file)
        return send_file(file, as_attachment = True,
                         download_name = fname,
                         mimetype = XLSX_MIMETYPE)
    else:
        return "Cannot generate!!"

if __name__ == '__main__':
    # This is used when running locally only. When deploying to Google App
    # Engine, a webserver process such as Gunicorn will serve the app. This
    # can be configured by adding an `entrypoint` to app.yaml.
    app.run(host='127.0.0.1', port=8080, debug=True)

