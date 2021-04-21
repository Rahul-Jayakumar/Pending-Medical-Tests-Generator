# Pending-Medical-Tests-Generator
The main reason why this program was created was to save the company paper. The company I work for is a private pathology company that performs numerous diagnostic tests. All these tests and their numerous parameters (such as correct sample types, sample viability, dates, time sample was taken, patient age, etc.) are viewed on a program known as Meditech known as a LIS (Laboratory Information System) and allows personnel to view all kinds of data relating to patients and tests logged for them. 


Meditech is a difficult program to work with as it is written in a proprietary language called MAGIC. This means, you can't change anything in the backend or use any sort of scripting function. The only way to get these pendings into a digital format would be via a pdf printer program like CutePDF or Windows PDF Print or a similar program while another program “pretends" to be an actual person typing in these pending test codes and other codes in a fairly linear sequence. I thought that there must be a way to emulate a person by getting the mouse and keyboard to work in the background. After some research, I came across the PyAutoGUI python automation module. It allows you to click or type things on your PC in the background without actually using a mouse or keyboard. It does this by use a pixel-by-pixel screen co-ordinate system to determine where buttons or locations are on the screen.


The code attached here is for a custom program written in Python that uses the PyAutoGUI automation module. 
