import lackey
from lackey import Screen, Pattern, Key, KeyModifier, Keyboard

l = lackey.Screen()
javascriptPrefix = 'javascript:(function(){'
postfix = '})();'


def pasteToAddressBar(cmd):
    global l
    l.wait(3)
    l.type('l', Key.CTRL)
    l.wait(2)
    l.type(javascriptPrefix)
    l.paste(cmd)
    l.paste(postfix)
    l.type(Key.ENTER)
    l.wait(3)
