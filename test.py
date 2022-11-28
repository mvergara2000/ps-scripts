from cryptography.fernet import Fernet
import json


dictionary = {
        "username": "username",
        "password": "password",
        "url": "url",
        "school": "school"
    }
json_object = json.dumps(dictionary, indent=4)
with open("test.json", "w") as outfile:
    outfile.write(json_object)    

key = Fernet.generate_key()
 
# string the key in a file
with open('filekey.key', 'wb') as filekey:
   filekey.write(key)


# opening the key
with open('filekey.key', 'rb') as filekey:
    key = filekey.read()
 
# using the generated key
fernet = Fernet(key)
 
# opening the original file to encrypt
with open('test.json', 'rb') as file:
    original = file.read()
     
# encrypting the file
encrypted = fernet.encrypt(original)
 
# opening the file in write mode and
# writing the encrypted data
with open('test.json', 'wb') as encrypted_file:
    encrypted_file.write(encrypted)
fernet = Fernet(key)
 
# opening the encrypted file
with open('test.json', 'rb') as enc_file:
    encrypted = enc_file.read()
 
# decrypting the file
decrypted = fernet.decrypt(encrypted)
 
# opening the file in write mode and
# writing the decrypted data
with open('test.json', 'wb') as dec_file:
    dec_file.write(decrypted)
f = open('test.json')
login= json.load(f)
x = login["username"]
print(x)