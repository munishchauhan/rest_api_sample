# print("Hello World")

from flask import Flask, jsonify, make_response, send_from_directory
from flask_swagger_ui import get_swaggerui_blueprint
from routes import request_api
import pypyodbc as odbc #pip install pypyodbc
import pandas as pd

DRIVER_NAME = 'SQL SERVER'
SERVER_NAME = 'HP-Elitebook'
DATABASE_NAME = 'DBApp'

connection_string = f"""
    DRIVER={{{DRIVER_NAME}}};
    SERVER={SERVER_NAME};
    DATABASE={DATABASE_NAME};
    Trust_Connection=yes;
"""


app = Flask(__name__)


@app.route("/static/<path:path>")
def send_static(path):
    return send_from_directory('static', path)


SWAGGER_URL = '/swagger'
API_URL = '/static/swagger.json'
swaggerui_blueprint = get_swaggerui_blueprint(
    SWAGGER_URL,
    API_URL,
    config={
        'app_name': 'Python API Sample'
    }
)
app.register_blueprint(swaggerui_blueprint, url_prefix=SWAGGER_URL)


app.register_blueprint(request_api.get_blueprint())


@app.route("/")
def hello_world():
    return "<p>Hello, World!</p>"


@app.route('/api/users/<string:name>')
def get_users(name):

    # --- Step 1: Excel Load ---
    df = pd.read_excel('Import Items Template.xlsx', engine='openpyxl')
    print(df.head())

    conn = odbc.connect(connection_string)
    cursor = conn.cursor()

    #--- Step 3: Insert Rows ---
    for index, row in df.iterrows():
        cursor.execute("""
            INSERT INTO tbCountryList (VCountryName, CurrencySymbol, CurrencyName)
            VALUES (?, ?, ?)
        """, (
            str(row["Item code"]) if not pd.isna(row["Item code"]) else None,
            str(row["Category"]) if not pd.isna(row["Category"]) else None,
            str(row["HSN"]) if not pd.isna(row["HSN"]) else None
        ))

    conn.commit()
    cursor.close()
    conn.close()

    # cursor.execute("SELECT * FROM tbCountryList")
    # list = cursor.fetchall()   

    # conn.close()

    # return jsonify(list)

    # return jsonify({
    #     "name": f"Hello {name}"
    # })
    return jsonify({"status": "success", "message": "Excel data imported"}) 

if __name__ == "__main__":
    app.run(debug=True)

