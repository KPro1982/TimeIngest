# import anvil.server

def change_name():
    anvil.server.connect("server_7UDURRX3SBYJOCU3HKUJCCUV-CWQSJE54273JU7WL")
    text = anvil.server.call('ChangeName', 'clientdanny')
    print("local result", text)