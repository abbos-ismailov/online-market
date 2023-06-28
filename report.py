lines = ['Readme', 'Python', "abbos", "zorku"]
with open('readme.txt', 'w') as f:
    for line in lines:
        f.write(line)
        f.write('\n')


# Importing library
import qrcode
 
# Data to be encoded
data = 'https://github.com/abbos-ismailov/online-market/blob/master/test_uchun.txt'
 
# Encoding data using make() function
img = qrcode.make(data)
 
# Saving as an image file
img.save('MyQRCode1.png')