
from datetime import date
from flask import jsonify, request, current_app, Blueprint
from flask import render_template, session, request, redirect, url_for, jsonify, send_from_directory 

import msal
import pandas as pd
import json
import os
from dotenv import load_dotenv
load_dotenv()

import checkLogged
import requests

from my_app import database, db, sharepoint_path

#########################################################################################################
## Gloval variables  
#########################################################################################################

eleaveDtl = db["eleave_dtl"]

#########################################################################################################
## BluePrint Declaration  
#########################################################################################################

entry = Blueprint('entry', __name__)

#########################################################################################################
## login
#########################################################################################################

def _load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        os.environ['CLIENT_ID'], authority=authority or os.environ['AUTHORITY'],
        client_credential=os.environ['CLIENT_SECRET'], token_cache=cache)

def _build_auth_code_flow(authority=None, scopes=None, redirect_uri=None):
    return _build_msal_app(authority=authority).initiate_auth_code_flow(
        scopes or [],
        redirect_uri or [])
        #redirect_uri=url_for("authorized2", _external=True))

def _get_token_from_cache(scope=None):
    cache = _load_cache()  # This web app maintains one cache per session
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if accounts:  # So all account(s) belong to the current signed-in user
        result = cca.acquire_token_silent(scope, account=accounts[0])
        _save_cache(cache)
        return result


@entry.route("/api/getPhoto/<email>")
@checkLogged.check_logged
def getPhoto(email=None):

    try:
    
    # if (os.environ['ENVIRONMENT']=="PROD"):                
    #     #check whether it's Macys's email account
    #     if  "@macys.com" not in session['email'].lower():
    #          return send_from_directory("frontend/public/static/img", "anonymous.jpg")        
    # else:
    #     print("Getting photo for development")       
    #     return send_from_directory("frontend/public/static/img", "anonymous.jpg")
    
        token = _get_token_from_cache(json.loads(os.environ['SCOPE']))
        if not token and not os.environ:
            return redirect(url_for("entry.login"))
    
        ## Getting photo          
        
        ##endpoint = "https://graph.microsoft.com/v1.0/me/photos/120x120/$value"

        endpoint = f"https://graph.microsoft.com/v1.0/users/{email}/photos/120x120/$value"
        ##endpoint = "https://graph.microsoft.com/v1.0/users/ken.yip@macys.com/photos/120x120/$value"
                
        photo_response = requests.get(  # Use token to call downstream service
            endpoint,
            headers={'Authorization': 'Bearer ' + token['access_token']},
            stream=True) 
        photo_status_code = photo_response.status_code
        if photo_status_code == 200:
            photo = photo_response.raw.read()
            return photo 
        else:        
            return  send_from_directory("frontend/build/static/img", "anonymous.jpg")
    except:
            return  send_from_directory("frontend/build/static/img", "anonymous.jpg")

      

@entry.route("/")
@checkLogged.check_logged
def index():    
    if not session.get("user"):
        return redirect(url_for("entry.login"))
    return render_template('index.html', user=session["user"], version=msal.__version__)

@entry.route("/login", defaults={'timeout':None}) 
@entry.route("/login/<timeout>") 
def login(timeout):
    if (timeout):
        print ("Entering login process with "  + timeout)
    # Technically we could use empty list [] as scopes to do just sign in,
    # here we choose to also collect end user consent upfront
 
    session["flow"] = _build_auth_code_flow(scopes=json.loads(os.environ['SCOPE']), redirect_uri=url_for("entry.authorized", _external=True))    
    session["flow2"] = _build_auth_code_flow(scopes=json.loads(os.environ['SCOPE']), redirect_uri=url_for("entry.authorized2", _external=True))    
    #  auth_uri an be added with prompt=login to force sign in     

    return render_template("login.html", auth_url=session["flow"]["auth_uri"], auth_url2=session["flow2"]["auth_uri"], version=msal.__version__, timeout_message=timeout)

@entry.route(os.environ['REDIRECT_PATH'])  # Its absolute URL must match your app's redirect_uri set in AAD
def authorized():
    try:
        print("Entering " + os.environ['REDIRECT_PATH'])
        cache = _load_cache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")
        # Vincent added below:
        #print ("email", json.dumps(result.get("id_token_claims")))
        #print ("email", result.get("id_token_claims").get('email'))
        session["email"] = (result.get("id_token_claims").get('email')).lower()      
        _save_cache(cache)
    except ValueError:  # Usually caused by CSRF
        pass  # Simply ignore them
        return render_template("auth_error.html", result={"error" : "Value Error", "error_description":"Not signed in yet !!"})    
    return redirect(url_for("entry.index"))

@entry.route("/#/ApprovalCenter")  # Its absolute URL must match your app's redirect_uri set in AAD
def authorized2():
    try:    
        cache = _load_cache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow2", {}), request.args)
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")
        # Vincent added below:
        #print ("email", json.dumps(result.get("id_token_claims")))
        #print ("email", result.get("id_token_claims").get('email'))
        session["email"] = (result.get("id_token_claims").get('email')).lower()      
        _save_cache(cache)
    except ValueError:  # Usually caused by CSRF
        pass  # Simply ignore them
        return render_template("auth_error.html", result={"error" : "Value Error", "error_description":"Not signed in yet !!"})
    return redirect(os.environ['APPROVAL_CENTER'])
    

@entry.route("/logout")
def logout():
    session.clear()  # Wipe out user and its token cache from session
    return redirect(  # Also logout from your tenant's web session
        os.environ['AUTHORITY'] + "/oauth2/v2.0/logout" +
        "?post_logout_redirect_uri=" + url_for("entry.index", _external=True))


@entry.route("/graphcall")
@checkLogged.check_logged
def graphcall():
    token = _get_token_from_cache(json.loads(os.environ['SCOPE']))
    if not token:
        return redirect(url_for("entry.login"))
    graph_data = requests.get(  # Use token to call downstream service
        os.environ['ENDPOINT'],
        headers={'Authorization': 'Bearer ' + token['access_token']},
        ).json()
    return render_template('display.html', result=graph_data)



## below for Reacj JS

def getTodayDate():
    return date.today().strftime("%m/%d/%y")  ## get today's date 


@entry.route('/api/getUserProfile',methods=['POST'])
@checkLogged.check_logged
def getUserProfile():                
    content = request.get_json() #python data          
    impersonatedUser = ""
    if content:
        impersonatedUser = content['impersonatedUser']        

    sessionData, status_code = establishSessionData(impersonatedUser)

    try:
        if (status_code == 200):
            return  jsonify(sessionData), status_code 
        else:
            return jsonify({'error_message' : 'Cannot find your RACF ID.  Please contact regional PBT !'}), status_code     
    except:       
        return jsonify({'error_message' : 'Cannot get your profile.  Please contact regional PBT !'}), status_code     


@entry.route('/api/impersonateUser',methods=['POST'])
@checkLogged.check_logged
def getImpersonateUser():            
    
  ## Getting racf ID and employee detail via racf 
  
    try:
        content = request.get_json() #python data     
        impersonatedUser = content['impersonatedUser']

        ## look for employee details via Mongo DB             
        sessionData, status_code = establishSessionData(impersonatedUser)      

        return  sessionData, status_code

    except Exception as e:     
        return "User impersonation error", status_code 

def construct_event_detail(event_name, **event_details):
    request_body = {
        'subject': event_name
    }
    for key, val in event_details.items():
        request_body[key] = val
    return request_body

    
   
def establishSessionData(impersonatedUser=""):

    try:        

        racf = ""

        if (os.environ['ENVIRONMENT']=="HEROKU"):            
            
            endpoint = "https://graph.microsoft.com/beta/me"                    
            token = _get_token_from_cache(json.loads(os.environ['SCOPE']))

            if not token and not os.environ:
                return redirect(url_for("entry.login"))

            racf_response = requests.get(  # Use token to call downstream service
                endpoint,
                headers={'Authorization': 'Bearer ' + token['access_token']}, stream=True
                ) 
            status_code = racf_response.status_code            

            ## onPremisesSamAccountName stores RACF ID - you can use MS Graph and endpoint to see                  

            if status_code == 200:
                pass
                racf_data =  racf_response.json()                                   
                racf = racf_data["onPremisesSamAccountName"].upper()                
            else:                 
                raise Exception("RACF ID failed to validate in the Active Directory.  Please contact regional PBT for assistance!")                                
             
        else:
            racf = current_app.config['APP_RACF'].upper()

        if len(impersonatedUser) > 0:            
            racf  = impersonatedUser    
        
        sessionData={}     

        ## look for employee details via Mongo DB        
        query =  { "staff.racf": racf}
        results = eleaveDtl.find_one(query)             
       
        years_str = os.environ['YEARS']      
        years = eval(years_str)
        years = pd.DataFrame(data=years)
        years.sort_values(by=["year"], ascending=True, inplace=True)
        years = years.to_json(orient="columns")        
        
        sessionData["userProfile"] = { 
            "email": results["staff"]["email"], 
            "userName" : results["staff"]["name"], 
            "office" : results["staff"]["office"], 
            "entitlement" : results['entitlement'],
            "racf" :  results["staff"]["racf"], 
            "superUser": results['staff']["superUser"] if len(impersonatedUser) ==0 else session["superUser"], 
            "powerBI": results["staff"]["powerBI"], 
            "environment":  os.environ["ENVIRONMENT"],   
            "databaseSchema":  "dev" if database[:3].lower() == "dev" else "prod",
            "staff": results["staff"],
            "years" : years,
            "switchUser" : os.environ["SWITCH_USER"]
        }            
       
        session['racf'] = sessionData["userProfile"]["racf"]        
        session['superUser'] = sessionData["userProfile"]["superUser"]        
        sessionData["sharePointPath"] = sharepoint_path
    
        return sessionData, 200 

    except Exception as e:             
        return "Session establish error !!", 501
             
               
  