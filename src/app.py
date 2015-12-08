from flask import Flask
from flask import render_template
from flask import request
from flask import redirect
from flask import url_for
from flask import make_response

from converter import convert_xlsx_to_vcard


app = Flask(__name__)
app.debug = True


@app.route('/', methods=['GET'])
def home():
    return render_template('index.html')


@app.route('/convert/', methods=['POST'])
def convert():
    contacts_file = request.files['contacts_file']
    start_row = int(request.form['start_row'])
    first_name_column_no = int(request.form['first_name_column_no'])
    last_name_column_no = int(request.form['last_name_column_no'])
    contact_no_column_no = int(request.form['contact_no_column_no'])
    column_map = {
        'first_name_column_no': first_name_column_no,
        'last_name_column_no': last_name_column_no,
        'contact_no_column_no': contact_no_column_no
    }
    contacts_vcf = convert_xlsx_to_vcard(contacts_file, column_map, start_row)
    response = make_response(contacts_vcf)
    response.headers['Content-Disposition'] = 'attachment; filename=contacts.vcf'
    return response


if __name__ == '__main__':
    app.run()
