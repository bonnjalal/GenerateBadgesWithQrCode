import os
import shutil
from openpyxl import load_workbook
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import qrcode



src_dir = os.getcwd() #get the current working dir
def copyAndRename(badge):
    # print(src_dir)

    # create a dir where we want to copy and rename
    # dest_dir = os.mkdir('Badges')
    # os.listdir()

    dest_dir = src_dir+"/Badges"
    src_file = os.path.join(src_dir, badge)
    shutil.copy(src_file,dest_dir) #copy the file to destination dir

    dst_file = os.path.join(dest_dir,badge)
    new_dst_file_name = os.path.join(dest_dir, 'name second name.png')

    os.rename(dst_file, new_dst_file_name)#rename
    os.chdir(dest_dir)

    return new_dst_file_name

def writeTxtOnBadge(name, affiliation,aff2, badge_img):
    # file = copyAndRename()
    
    im = Image.open(badge_img)
    W, H = im.size

    # x = W/10
    # y1 = H/2.5

    # Call draw Method to add 2D graphics in an image
    I1 = ImageDraw.Draw(im)
     # Custom font style and font size
    myFont = ImageFont.truetype(src_dir + '/montesrrat-bold.ttf', 72)
    myFont2 = ImageFont.truetype(src_dir + '/montesrrat-bold.ttf', 55)
     
    # Add Text to an image
    # I1.text((x, y), "Nice Car", font=myFont, fill =(255, 255, 255))
    _, _, w, h = I1.textbbox((20, 0), name, font=myFont)
    I1.text(((W-w)/2, (H-h)/2.3), name, font=myFont, fill=(255,255,255))

    _, _, w2, h2 = I1.textbbox((20, 0), affiliation, font=myFont2)
    I1.text(((W-w2)/2, (H-h2)/1.85), affiliation, font=myFont2, fill=(255,255,255))

    _, _, w3, h3 = I1.textbbox((20, 0), aff2, font=myFont2)
    I1.text(((W-w3)/2, (H-h3)/1.69), aff2, font=myFont2, fill=(255,255,255))

    # Display edited image
    # im.show()
     
    # Save the edited image
    badgePath = src_dir+ "/Badges/" + name + '.png'
    im.save(badgePath)
    return badgePath

def generateQr(data:str):
    # Creating an instance of QRCode class
    qr = qrcode.QRCode(version = 1,
                       box_size = 15,
                       border = 0)
    
    # Adding data to the instance 'qr'
    qr.add_data(data)
    
    qr.make(fit = True)
    img = qr.make_image(fill_color = '#1c2813',
    # img = qr.make_image(fill_color = 'black',
                        back_color = 'white')
    
    qrPath = src_dir + '/Badges/QR_' +data+'.png'
    img.save(qrPath)

    return qrPath

def putQrOnBadge(qrPath, badgePath):
    # Opening the primary image (used in background) 
    img1 = Image.open(badgePath) 
    W, H = img1.size
      
    # Opening the secondary image (overlay image) 
    img2 = Image.open(qrPath) 
    w, h = img2.size 
    # Pasting img2 image on top of img1  
    # starting at coordinates (0, 0) 
    w1 = w  + w/7
    h1 = h + w/10
    img1.paste(img2, (int(W-w1),int(H-h1)))
      
    # Displaying the image 
    # img1.show()

    img1.save(badgePath)

wb = load_workbook('Badges_Names.xlsx')
ws = wb["Names in Badge"]

# Name in Badge (B)
# Affiliation (C)
# Role (D)
for i in range(2,127):
    name = ws.cell(row = i, column=2).value
    aff = ws.cell(row = i, column=3).value
    aff2 = ws.cell(row = i, column=4).value
    role = ws.cell(row=i, column=5).value 

    if aff2 == None:
        aff2 = ""
    print(name)

    back = src_dir + "/" + str(role) + ".png"
    badge_path = writeTxtOnBadge(name, aff,aff2,back)
    qr_path = generateQr(str(name))
    putQrOnBadge(qr_path, badge_path)
    os.remove(qr_path)


