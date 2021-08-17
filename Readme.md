# Attandance Analyzer

## Attandance managment using Face Recognition and  OpenCV in Python

Prerequisites
1.Python
2.Numpy
3.OpenCV
4.pyxl

### Installing 

Install Numpy 
pip install numpy

Install OpenCV via anaconda:
pip install opencv

install the other modules using >>pip install name of the module


### Running the tests


1.Place Images for training the classifier in trainingImages folder.If you want to train clasifier to recognize multiple people then add each persons folder in separate label markes as 0,1,2,etc and then add their names along with labels 

2.Run tester.py first since it generates training.yml file that is being used in "videoTester.py" script.

3.Place some test images in TestImages folder that you want to predict  in tester.py file(To generate test images for training classifier use videoimg.py file)

4.To do test run via tester.py give the path of image in test_img variable

5.Use "videoTester.py" script for predicting faces realtime via your webcam.

*follow the above steps and test the code for accuracy.

6.Finally run main.py to mark the attandance.

