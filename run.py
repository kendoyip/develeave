from my_app import create_app  # from the app package __init__
import os

app =  create_app()

# For local run only

if __name__ == '__main__':
    flask_app = create_app()
    with flask_app.app_context():        
        pass     
    flask_app.run(host='0.0.0.0', port=5000, debug=True)