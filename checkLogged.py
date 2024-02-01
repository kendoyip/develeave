from flask import session, request, redirect, url_for, jsonify 
import os
# decorator to check whether the session is still valid 

def check_logged(f):
    def wrapper(*args, **kwargs):        
        if (os.environ['ENVIRONMENT']=="HEROKU" or os.environ['ENVIRONMENT']=="HEROKU"):
            #print("running check_logged to check whether it is logged")
            if not (request.headers.get('api')) and not session.get("user"):
                return redirect(url_for("entry.login"))
            if (request.headers.get('api')) and not session.get("user"):                   
                return jsonify({"error_message": "You session seems to have timed out.  Please logout and login again !! "}), 599                           
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
    return wrapper