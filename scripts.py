"""changeGL.py
Automate remapping a FIMS GL Account code via
Tools -> System Utilities -> Admin Utilities -> Finance Utilities -> Change Natural Account
"""

import pyautogui as bot
import pandas as pd
import xlrd
import time

#import yrdata

bot.FAILSAFE = False # disables the fail-safe

screen_res = str(bot.size()[0]) + 'x' + str(bot.size()[1])

# Dict of dicts, storing (x,y) location of menu item for a given screen
# resolution (assumes FIMS is running at full screen).
# Format is: {'MENU LOCATION': {ScreenResolution: (X,Y)}}
menu_locations = {'tools': {'1366x768': (420,32)},
                  'system_utils': {'1366x768': (512,166)},
                  'admin_utils': {'1366x768': (805,580)},
                  'fin_utils': {'1366x768': (1016,661)},
                  'change_net_account': {'1366x768': (764, 482)},
                  }


def click_item(menu_item, screen_res = screen_res, menu_locations=menu_locations, logging=False):
    """
    Given a menu item name, click the mouse at the x,y coords
    :param menu_item:
    :return:
    """
    x,y = menu_locations[menu_item][screen_res]
    bot.click(x,y)
    time.sleep(1)
    if logging: print('clicked ' + menu_item)

    return True


def map_gl(fisc_yr, gl_old, gl_new, logging=False):

    time.sleep(1)

    click_item('tools')
    click_item('system_utils')
    click_item('admin_utils')
    click_item('fin_utils')
    click_item('change_net_account')

    time.sleep(1)

    # Click the "OK" button on the message box
    bot.press('enter')

    time.sleep(1)

    # Enter data in "Global Change for GL Natural Account" window
    # steps:    1) Enter "Old" Account number; /tab
    #           2) Enter "New" Account number; /tab*3
    #           3) Select FY: Type number "2" n times (once = 2000, twice=2001...)
    #           4) /tab*3, type "Enter" (Selects "OK" button)

    bot.typewrite(gl_old)
    bot.press('tab')

    time.sleep(1)

    bot.typewrite(gl_new)
    # Tab 3 times to select GL Year option
    bot.press('tab')
    bot.press('tab')
    bot.press('tab')

    time.sleep(1)

    # Enter "2" enough times to cycle to desired fiscal yr
    number_of_2s = int(fisc_yr) - 1998
    for _ in range(number_of_2s):
        bot.press('2', pause=.1)

    time.sleep(1)

    # Tab 3 times to select "Proceed" button
    bot.press('tab')
    bot.press('tab')
    bot.press('tab')

    # "Proceed" button is highlighted; hit "enter" to run process
    bot.press('enter')

    err_count = 0
    print(f'Beginning conversion of gl account {gl_old} to {gl_new} for {fisc_yr}')

    # Process will run for a long time. Check to see when completion message box
    # "Global Change for GL Natural Account" window appears - then click "Enter"
    for i in range(1000):

        if logging and i%10 == 0:
            print('Cycle Number ' + str(i))

        time.sleep(10)

        # Check to see if an error box pops up; click "OK" if it does
        gl_error_box_xy = \
            bot.locateOnScreen('img/' +
                               'gl_err2.png',
                               region=(200,200, 1000,600))

        if gl_error_box_xy is not None:
            print('Error found converting GL Account ' + gl_old)
            bot.screenshot('img/errors/' + gl_old + '_' + str(err_count) + '.png')
            err_count += 1
            bot.press('enter')

        done_msg_box_xy = \
            bot.locateOnScreen('img/message_global_change_done_v2_1366x768.png',
                               region=(450,290, 890,475))  # region=(473,307, 888,467))

        if done_msg_box_xy is not None: #box appeared
            bot.press('enter')  # Selects "OK" button on window
            print(f'Completed conversion of GL Account {gl_old} to {gl_new} for FY {fisc_yr}')
            break # Exit for loop


def screengrab(filename):
    time.sleep(5)

    region = bot.size() # (0,0, 1920,1080)

    bot.screenshot(filename + '.png')
    return True


def get_xy():
    for i in range(1000):
        time.sleep(.1)
        x,y = bot.position()
        print(str(x) + ', ' + str(y))

    return True
