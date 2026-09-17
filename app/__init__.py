import os
from flask import Flask
from dotenv import load_dotenv

load_dotenv()

def create_app():
    app = Flask(__name__, static_folder='static', template_folder='templates')
    app.secret_key = os.environ.get('SECRET_KEY', 'vendas_secret_key_postgres_18_chave_segura')
    app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
    
    from app.routes import main_bp
    app.register_blueprint(main_bp)

    return app

