# Pending-Medical-Tests-Generator
The main reason why this program was created was to save paper. The company I work at is a private pathology company that performs numerous diagnostic tests. All these tests and their numerous parameters (such as correct sample types, sample viability, dates, time sample was taken, patient age, etc.) are viewed on a program known as Meditech which is a program written in a proprietary language known as MAGIC. 
This program is known as a LIS (Laboratory Information System) and allows personnel to view all kinds of data relating to patients and tests logged for them. Since we were dealing with patient samples which have a viability period, we as medial scientists have to constantly track and query sample locations throughout the country a few times a day.
There are numerous scientists and technicians running a large variety of tests and all of these tests have numerous test codes which are their identifiers on Meditech. To see all the patients logged for each different test, one has to type the test code out in a specific section of the program known as the "Pendings" page. Here the codes need to be entered along with different test site codes. These records are then printed on paper. The lab has Pending’s about 30 employee and each one prints pendings around 3 times a day. Each pending set is usually around a minimum of 30 pages. It can go up to around 90 pages depending on how busy it is. The amount of paper used adds up very quickly along with heavy cost implications.
Meditech is a difficult program to work with as it is written in a proprietary language called MAGIC. This means, you can't change anything in the backend or use any sort of scripting function. The only way to get these pendings into a digital format would be via a pdf printer program like CutePDF or Windows PDF Print or a similar program while another program “pretends" to be an actual person typing in these pending test codes as well as test sites and clicks a whole bunch of buttons in a fairly linear sequence.
I thought that there must be a way to emulate a person by getting the mouse and keyboard to work in the background. After some research, I came across the amazing PyAutoGUI python automation module written by Al Sweigart. This allows you to click or type things on your PC in the background without actually using a mouse or keyboard. It does this by use a pixel-by-pixel screen cp-ordinate system to determine where buttons or locations are on the screen.
The code attached here is for a custom program written in Python that heavily utilizes the PyAutoGUI automation module. It is written for a 1280x1024 resolution screen as that is what is primarily used at the company. The PyAutoGUI values can be adjusted for any screen with some tweaking of the source code. There are comments in the .py file to explain everything. This README is essentially just a problem statement with some background on why the idea of this program was conceived.
