# from flask import Flask, render_template

# app = Flask(__name__, static_folder='static', template_folder='static')

# @app.route('/')
# def index():
#     return render_template('index.html')

# @app.route('/print')
# def print_view():
#     return render_template('print.html')

# if __name__ == '__main__':
#     app.run(debug=True)


from flask import Flask, render_template, send_from_directory

app = Flask(__name__, static_folder='static', template_folder='static')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/print')
def print_view():
    return render_template('print.html')

# Add this route to handle static files
@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory(app.static_folder, filename)

if __name__ == '__main__':
    app.run(debug=True)