from flask import Flask, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from workreport import getAllData, getAllID, findData, writeData

# If `entrypoint` is not defined in app.yaml, App Engine will look for an app
# called `app` in `main.py`.
app = Flask(__name__)


@app.route('/')
def hello():
    """Return a friendly HTTP greeting."""
    alldata = getAllData('102Fbm3HKFOeCg6Q0CEAQE88CLQ3xvtW3bTPDGydx02A',"2021前期")
    ids = getAllID(alldata)
    print(ids)
    ret = "Workreport generator for project TA 2021<br/>\n"
    ret += '<form action="/create">\n'
    ret += 'id: <select name="id">\n'
    for i in ids:
        ret += '<option value="{0}">{1}\n'.format(i,i)
    ret += '</select><br/>\n'
    months=[5,6,7]
    ret += 'month: <select name="month">\n'
    for m in months:
        ret += '<option value="{0}">{1}\n'.format(m,m)
    ret += '</select><br/>'
    ret += '<input type="submit">\n'
    ret += '</form>\n'
    return ret

@app.route('/create')
def create():
    stuid = request.args['id']
    month = request.args['month']
    alldata = getAllData('102Fbm3HKFOeCg6Q0CEAQE88CLQ3xvtW3bTPDGydx02A',"2021前期")
    data = findData(stuid, int(month), alldata)
    file = writeData(stuid, month, data)
    XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return send_file(file, as_attachment = True,
                     download_name = file,
                     mimetype = XLSX_MIMETYPE)

if __name__ == '__main__':
    # This is used when running locally only. When deploying to Google App
    # Engine, a webserver process such as Gunicorn will serve the app. This
    # can be configured by adding an `entrypoint` to app.yaml.
    app.run(host='127.0.0.1', port=8080, debug=True)

