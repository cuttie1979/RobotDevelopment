## Installation ##

Install Visual Studio Code from the following link: https://code.visualstudio.com/
Install Anaconda3 from the following link: https://www.anaconda.com/

Once installed we need to open a command prompt (search for cmd on Windows or Terminal on Mac)
Create environment for your robot with the following line:
    conda create --name robotdevelopment python=3.10

    Answer yes from the question

Once it completed we can use it in Visual Studio Code:
CTRL+SHIFT+P
Select python interpreter, and choose the newly created environment. For each robot it would be good to use separate environment

We use two files for the development, one is requirements.txt where we list all of the used frameworks which helps our development
Next is the robot program file e.g. robot.py what is the robot 

Documentation is our good friend, it helps to understand how different functions work:
https://robocorp.com/docs/development-guide/excel
https://rpaframework.org/libraries/desktop/

We will use in this example the robocorp framework, but based on the tasks we can use more

## Usefull extensions ##
PyPI Assistant for VS Code