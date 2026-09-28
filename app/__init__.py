import os
from flask import Flask
from dotenv import load_dotenv

load_dotenv()

def create_app():
    app = Flask(__name__, static_folder='static', template_folder='templates')
    app.secret_key = os.environ.get('SECRET_KEY', 'vendas_secret_key_postgres_18_chave_segura')
    app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
    
    @app.after_request
    def add_header(response):
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '-1'
        return response

    from app.routes import main_bp
    app.register_blueprint(main_bp)

    return app

