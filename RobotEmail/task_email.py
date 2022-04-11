'''
#############################
# Program: Robot send email #
# Author: Popovics Laszlo   #
#############################
'''

from RPA.Desktop import Desktop
import time

desktop = Desktop()

desktop.click(locator="point:240,1070")
time.sleep(1)
desktop.press_keys('cmd','n')
time.sleep(1)
desktop.type_text('lpopovics@deltecbank.com')
desktop.press_keys('tab')
desktop.type_text('Test email')
desktop.press_keys('tab')
desktop.press_keys('cmd','a')
desktop.type_text('This is a test email from Laszlo\'s robot\n' +
                  '\n' + 
                  'Have a great day!\n' + 
                  '\n' +
                  '        Laszlo')
desktop.press_keys('cmd','e')
time.sleep(1)
desktop.click(locator="point:550,510")
time.sleep(1)
desktop.click(locator="point:690,420", action='double_click')
time.sleep(1)
desktop.click(locator="point:715,400", action='double_click')
time.sleep(1)
desktop.click(locator="point:740,440", action='double_click')
desktop.press_keys('cmd','enter')