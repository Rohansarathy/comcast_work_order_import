import json

credentials = {
    "FuseURL": "https://fuse.i-t-g.net/login.php",
    "Fusername": "ybot",
    "Fpassword": "Bluebird1@3",
    # "Main_file": r"C:\Users\Administrator\OneDrive - ITG Communications, LLC\Work Order Import\comcast\NewDynamicInput.xlsx",
    "Main_file": r"C:\Users\Administrator\Documents\Work Order Import\comcast\NewDynamicInput.xlsx",
    "extaction_file": r"C:\Users\Administrator\OneDrive - ITG Communications, LLC\Work Order Import\comcast\Comcast FUSE uploads mapping.xlsx",
    "Logfile" : r"C:\Users\Administrator\OneDrive - ITG Communications, LLC\Work Order Import\comcast\comcast_logs",  
    "ybotID":"ybot@itgext.com",
    "sdavis":"sdavis@itgcomm.com"
}

with open("work_order.json", "w") as json_file:
    json.dump(credentials, json_file)