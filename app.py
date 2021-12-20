from flask import Flask, render_template
import pandas as pd
import stock


stock.stock_data()

app = Flask(__name__)

@app.route('/')


def table():
	data = pd.read_excel('Stock-Book.xlsx')
	return render_template('table.html', tables=[data.to_html()], titles=[''])


if __name__ == "__main__":
	app.run(host="localhost", port=int("5000"))
