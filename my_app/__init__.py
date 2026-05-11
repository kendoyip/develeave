# -*- coding: utf-8 -*-
from flask import Flask, session
from flask_session import Session  # https://pythonhosted.org/Flask-Session

import json
from datetime import timedelta
import os 
from dotenv import load_dotenv
load_dotenv()

from pymongo import MongoClient
import certifi

#### Mail function 
from flask_mail import Mail

# Below for SahrePoint
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
# from office365.sharepoint.files.file import File
# from office365.sharepoint.listitems.caml.caml_query import CamlQuery  
# from office365.runtime.http.request_options import RequestOptions
# from office365.sharepoint.files.file_creation_information import FileCreationInformation

env = os.environ['ENVIRONMENT']

client = MongoClient(os.environ['MONGODB_URL'], tls=True, tlsAllowInvalidCertificates=True, tlsCAFile=certifi.where(),  maxPoolSize=100)
database = os.environ['DEV_DATABASE'] if env == "LOCAL" else os.environ['DATABASE']

# Use manual modified lifetime
try:
    session_collection = client[os.environ['DEV_DATABASE'] if env == "LOCAL" else os.environ['DATABASE']]['sessions'] 
    session_collection.drop_index("expiration_1")
except:
    pass


mail = Mail()
db = client[database]

site_url = 'https://macysinc.sharepoint.com/sites/MMGOverseas/'
app_principal = {
     'client_id': os.environ['SHAREPOINT_CLIENT_ID'],
     'client_secret': os.environ['SHAREPOINT_CLIENT_SECRET'],
}

sharepoint_path =  os.environ['SHAREPOINT_PATH']

context_auth = AuthenticationContext(url=site_url)
context_auth.acquire_token_for_app(client_id=app_principal['client_id'], client_secret=app_principal['client_secret'])    
ctx = ClientContext(site_url, context_auth)


def create_app():

    static_folder = "frontend/build/static"
    template_folder = "frontend/build"   
        
    app = Flask(__name__, static_folder=static_folder, template_folder=template_folder)   


    #if env == "HEROKU":        
        # read MailerToGo env vars
        #app.mailertogo_host     = os.environ.get('MAILERTOGO_SMTP_HOST')
        #app.mailertogo_port     = os.environ.get('MAILERTOGO_SMTP_PORT', 587)
        #app.mailertogo_user     = os.environ.get('MAILERTOGO_SMTP_USER')
        #app.mailertogo_password = os.environ.get('MAILERTOGO_SMTP_PASSWORD')
        #app.mailertogo_domain   = os.environ.get('MAILERTOGO_DOMAIN', "mydomain.com")
        #app.recipient_domain = os.environ.get('RECIPIENT_DOMAIN')
        #app.macys_domain = os.environ.get('MACYS_DOMAIN')


    if env == "LOCAL":
        # read Local .env
    #    app.mailertogo_host     = os.environ["MAILERTOGO_SMTP_HOST"]
    #    app.mailertogo_port     = os.environ["MAILERTOGO_SMTP_PORT"]
    #    app.mailertogo_domain   = os.environ["MAILERTOGO_DOMAIN"]
    #    app.config['recipient_domain'] = os.environ["RECIPIENT_DOMAIN"]
        app.config['APP_EMAIL']= os.environ['APP_EMAIL']
        app.config['APP_RACF']= os.environ['APP_RACF']
      
    
    #app.config['MAIL_SERVER']='smtp.us-west-1.mailertogo.net'
    #app.config['MAIL_PORT'] = 587
    #app.config['MAIL_SERVER']='appmailos.federated.fds'
    #app.config['MAIL_PORT'] = 25
    app.config['SECRET_KEY']= os.environ['SECRET_KEY']
    
    mail.init_app(app)       

    app.config['SESSION_TYPE'] = 'mongodb'  
    app.config['SESSION_KEY_PREFIX'] = 'session:' 


    app.config['SESSION_MONGODB'] = client    
    app.config['SESSION_MONGODB_DB'] =  os.environ['DEV_DATABASE'] if env == "LOCAL" else os.environ['DATABASE']
    app.config['SESSION_MONGODB_COLLECT'] = 'sessions'
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes = int(os.environ['SESSION_TIMEOUT']))


    app.config['YEARS'] = os.environ['YEARS']
 
    Session(app)

    app.config['UPLOAD_FOLDER'] = '.'
    #app.config['MAX_CONTENT_LENGTH'] = 100 * 1024    # 100K       
    app.config['MAX_CONTENT_LENGTH'] = int(os.environ['UPLOAD_MAX_SIZE'])   

    # This section is needed for url_for("foo", _external=True) to automatically
    # generate http scheme when this sample is running on localhost,
    # and to generate https scheme when it is deployed behind reversed proxy.
    # See also https://flask.palletsprojects.com/en/1.0.x/deploying/wsgi-standalone/#proxy-setups
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

    from my_app.entry.views import _build_auth_code_flow   
    app.jinja_env.globals.update(_build_auth_code_flow=_build_auth_code_flow)  # Used in template

    from my_app.entry.views import entry
    from my_app.eleave.views import eleave
    from my_app.spoint.views import spoint
    #from my_app.maintenance.views import maintenance
    app.register_blueprint(entry, url_prefix='')
    app.register_blueprint(eleave, url_prefix='/eleave')
    app.register_blueprint(spoint, url_prefix='/spoint')
    #app.register_blueprint(maintenance, url_prefix='/maintenance')


    return app
