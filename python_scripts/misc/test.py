import os


try:
    with open("/Users/senthil/Desktop/Senthil/myTesting/test.xml", "r") as file:
        xml_file = file.read()
        print (xml_file)
except OSError:
    print("invalid path!")
    # pass