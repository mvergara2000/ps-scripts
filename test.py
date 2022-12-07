import os
from cryptography.fernet import Fernet
import json
def main():
    #with open('filekey.key', 'rb') as filekey:
    k = open('filekey.key')
    key = k.read()
    dictionary = {
        "username": "username",
        "password": "password",
        "url": "url",
        "school": "school"
        }
    json_object = json.dumps(dictionary, indent=4)
    #with open("test.json", "w") as outfile:
        #outfile.write(json_object) 
    #encrypt(key)
    decrypt(key)

def generateKey():
# key generation
    key = Fernet.generate_key()
 
# string the key in a file
    with open('filekey.key', 'wb') as filekey:
        filekey.write(key)
def encrypt(key):
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


def decrypt(key):
    # using the key
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

if __name__=="__main__":
    main()