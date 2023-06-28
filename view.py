import qrcode
def qrCode():
    # Data to be encoded
    data = 'QR Code using make() function'
    
    # Encoding data using make() function
    img = qrcode.make(data)
    
    # Saving as an image file
    img.save('QR Code products_baza.png')
    
    
    