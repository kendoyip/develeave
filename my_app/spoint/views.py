# -*- coding: utf-8 -*-

from flask import jsonify, request, Blueprint,  jsonify,  send_file
import os
from dotenv import load_dotenv
load_dotenv()
import checkLogged
from io import BytesIO

#########################################################################################################
## For SharePoint 
#########################################################################################################
from office365.sharepoint.files.file import File
from office365.sharepoint.listitems.caml.caml_query import CamlQuery  
#from office365.runtime.http.request_options import RequestOptions
from office365.sharepoint.files.file_creation_information import FileCreationInformation
from my_app import ctx

#########################################################################################################
## Gloval variables  
#########################################################################################################
web = ctx.web
ctx.load(web)
ctx.execute_query()

#########################################################################################################
## BluePrint Declaration  
#########################################################################################################

spoint = Blueprint('spoint', __name__)

#########################################################################################################
## Functions
#########################################################################################################
import math 
def convert_size(size_bytes):
   if size_bytes == 0:
       return "0B"
   size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
   i = int(math.floor(math.log(size_bytes, 1024)))
   p = math.pow(1024, i)
   s = round(size_bytes / p, 2)
   return "%s %s" % (s, size_name[i])

ALLOWED_EXTENSIONS = set(['pdf', 'png', 'jpg', 'jpeg', 'xlsx', 'doc', 'docx', 'ppt', 'pptx', 'zip'])

def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


#########################################################################################################
## SharePoint 
#########################################################################################################


# @spoint.route('/api/test',methods=['POST'])
# @checkLogged.check_logged
# def test():
#     print("Web site title test : {0}".format(web.properties['Title']))

    # get SharePoint site - https://macysinc.sharepoint.com/sites/OSO/_api/Web/siteusers

    # try: 
    #         _request = RequestOptions("{0}/_api/web/siteusers".format(os.environ['SHAREPOINT_SITE']))
    #         _response = ctx.execute_request_direct(_request)
    #         _content_all = json.loads(_response.content)
    #         siteUsers = []
    #         for item in _content_all['d']['results']: 
    #             _username = item['Title']
    #             _email = item['Email']
    #             _id = item['Id']
    #             siteUser = { 
    #                     "id":_id,                
    #                     "username": _username,
    #                     "email": _email               
    #                 } 
    #             siteUsers.append(siteUser)
    #             session["siteUsers"] = siteUsers     
    #             print('under test', session['siteUsers'])
        
    # except Exception as e:        
    #         return f"***Error getting site user list from SharePoint***, {str(e)}", 500
            

    # return 'OK', 200


###################################################################################
#  SharePoint Section - Start
####################################################################################

@spoint.route('/api/getsharepointfiles',methods=['POST'])
@checkLogged.check_logged
def getSharepointfiles():       

    try:        
        content = request.get_json() #python data     
                      
        folder = content['racf']
        sharePointID = content['sharePointID']        
        year =  content['year']        
        ## below are getting from environment 
        sharePointReport = os.environ['SHAREPOINT_REPORT'] + str(year)
        relative_url = os.environ['SHAREPOINT_PATH'] + str(year) +  "/" + folder        

                       
        libraryRoot = ctx.web.get_folder_by_server_relative_path(relative_url)
        ctx.load(libraryRoot)
        ctx.execute_query()


        #if you want to get the files in the folder        
        files = libraryRoot.files
        ctx.load(files)
        ctx.execute_query()

        for file in files:    
            _name = file.properties["Name"]    
            print("Folder {0}, File name: {1}".format(folder, _name))

        #if you want to get the items in the folder        
        caml_query = CamlQuery()
        # Need _x0020 for space if the field has a space in the word 
        caml_query.ViewXml = '''<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name='SharePointID' /><Value Type='Text'>{0}</Value></Eq></Where></Query></View>'''.format(sharePointID)
        caml_query.FolderServerRelativeUrl = relative_url
    
        # 3 Retrieve list items based on the CAML query         
        oList = ctx.web.lists.get_by_title(sharePointReport) 
        items = oList.get_items(caml_query) 
        ctx.execute_query()

        sharePoint_array = []         
        for item in items:                
            _sharePointID = item.properties["SharePointID"]                                
            _id  = item.properties["Id"]                
            list_item  = item.expand(["File"])            
            list_item = ctx.web.lists.get_by_title(sharePointReport).get_item_by_id(_id).expand(["File"])
            ctx.load(list_item)
            ctx.execute_query()         
            _size = list_item.file.properties['Length']                         
            _size = convert_size(int(_size))       
            
            
            sharePoint_items = { 
                "folder":folder,
                "filename": list_item.file.properties['Name'],
                "sharePoint_id": _sharePointID,                
                "size": _size,
                "url": "https://macysinc.sharepoint.com" + list_item.file.properties["ServerRelativeUrl"],
                "relative_path" : list_item.file.properties["ServerRelativeUrl"],
                "unique_id": list_item.file.unique_id                
            }          
            ##print('id', list_item.file.unique_id)
            sharePoint_array.append(sharePoint_items)
       
        return  jsonify(sharePoint_array), 200         

    except Exception as e:
        print('getSharePoint Exceptoin', e)
        return "Error getting SharePoint", 502    


#######  SharePoint delte photos from SharePoint
@spoint.route('/api/deleteSPfile',methods=['POST'])
@checkLogged.check_logged
def delete_sp_file():
        
    content = request.get_json() #python data     
    _id = content['_id'] # file id
    print('deleting ', _id)
    try:  
        f = ctx.web.get_file_by_id(_id)
        ##ctx.execute_query()
        f.delete_object()
        ctx.execute_query()        
        print('deleteing ok')
        return  "OK", 200 
       
    except Exception:        
        return "Error", 509

@spoint.route('/api/downloadSharePointFile', methods=['GET'])
@checkLogged.check_logged
def download_sharepoint_file():
    try:
      ##print('file ', request.headers['file-url']  )

      file_url = request.headers['file-url'] 
      file_url = file_url.replace("'","%27%27") ## replace character ' 
      file_url = file_url.replace("#","%23")    ## replace character #       
      
      _response = File.open_binary(ctx, file_url)                 
      data = BytesIO(_response.content)               
      return send_file(data, attachment_filename='whatever.jpg', mimetype='image/jpg')                                                                 


    except:
      return 'Error', 501



@spoint.route('/api/upload', methods=['POST'])
@checkLogged.check_logged
def upload_file():
        
  if request.method == 'POST':
              
            try:
                sharePointID = request.headers['sharePointID']                
                relative_url =  request.headers['relative_url']
                name =  request.headers['name']
                office =  request.headers['office']

                # Get what a list of files in the Manufacturer folder 
                mf_dir = ctx.web.get_folder_by_server_relative_path(relative_url)
                mf_files = mf_dir.files
                ctx.load(mf_files)
                ctx.execute_query()                                
    
                #print('Upload/Inspection_id is {0}, {1}, {2}'.format(inspection_id, su_no, mf_no))                   
                # check if the post request has the files part
                if 'files[]' not in request.files:
                        #flash('No file part')
                        return "No files", 406
              
                files = request.files.getlist('files[]')      
           
                newfiles = []

                for file in files:                           
                    
                    ##file.seek(0, os.SEEK_END)
                    ##file_length = file.tell()
                    ##print('file.length', file_length)
                    for mf_file in mf_files:
                        if  mf_file.properties["Name"] == file.filename:                            
                            #print('your file name has existed')
                            #return f"Your file name has existed.  Please upload another name! ", 501            
                            errMessage = f"Same file name has already existed. Please use other file name !"
                            return jsonify({"error_message" : errMessage }), 501                                            
                                          

                    if file and allowed_file(file.filename):                                
                        
                        mimetype = file.content_type
                        filename = file.filename    
                        
                        # 6/23/22 remove below as secure_filename removes non-ascii characters.
                        # filename = secure_filename(filename)                          

                        target_folder = ctx.web.get_folder_by_server_relative_path(relative_url)                        
                        ctx.execute_query()
                        
                        info = FileCreationInformation()                        
                        info.content = file.read()        
                        #enable below for control of each size of the file. 
                        # if  len(info.content) >  1024 * 1024 * 1:    
                        #     return f"The size of {file.filename} exceeds 2 MB limit !", 555                                                                            
                        
                        info.url = filename                                  
                        info.overwrite = True                                
                        upload_file = target_folder.files.add(info)             
                   
                        ctx.execute_query()
                        list_item = upload_file.listItemAllFields # get associated list item                                 
                        list_item.set_property("SharePointID", sharePointID)
                        list_item.set_property("StaffName", name)
                        list_item.set_property("OfficeName", office)                        
                        list_item.update()                                
                        ctx.execute_query()                                 
                       
                    else:
                        return f"Your file type is not allowed", 502

                return "OK",200
                                
            except Exception as e:                    
                return "Upload Error", 507


###################################################################################
#  SharePoint Section - End
####################################################################################
