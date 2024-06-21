# import cv2 as cv
# events = [i for i in dir(cv) if 'EVENT' in i]
# print( events )

# #%%
# def Capture_Event(event, x, y, flags, params):
# 	# If the left mouse button is pressed
# 	if event == cv.EVENT_LBUTTONDOWN:
# 		# Print the coordinate of the 
# 		# clicked point
# 		print(f"({x}, {y})")
		
# #%%
# if __name__=="__main__":
# 	# Read the Image.
# 	img = cv.imread("P:/Pertinent Files/Python/saved_images/distribution.png", 1)
# 	# Show the Image
# 	cv.imshow('image', img)
# 	# Set the Mouse Callback function, and call
# 	# the Capture_Event function.
# 	cv.setMouseCallback('image', Capture_Event)
# 	# Press any key to exit
# 	cv.waitKey(0)
# 	# Destroy all the windows
# 	cv.destroyAllWindows()
# #%%

# # path="P:/Pertinent Files/Python/saved_images/distribution.png"
# # img=Image.open(path)

# # # imgcropped=img.crop(box=(40,48,1200,1200))
# # imgcropped=img.crop(box=(75,75,1200,1200))

# # imgcropped.show()


#%%

import cv2

def main():
    image=cv2.imread("P:/Pertinent Files/Python/saved_images/distribution.png")
    dimensions=image.shape
    height=dimensions[0]
    width=dimensions[1]
    # 272,86 - 272,574 - 1141,574 - 1141,86
    height=574
    width=1141
    cropped_image=image[86:574,270:1141]
    cv2.imwrite("P:/Pertinent Files/Python/saved_images/croppedbird.png", cropped_image)

if __name__=='__main__':
    main()