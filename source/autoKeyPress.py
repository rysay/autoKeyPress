# coding: utf-8
import pyautogui
import time
import decimal


def main():
    """メイン処理
    """
    pyautogui.FAILSAFE = False
    text = """\\t   \\n   \\r
!   ”   #   $   %   &   '   (   )   =   ~   |
`   +   *   {   }   <   >   ?   _   yen
1   2   3   4   5   6   7   8   9   0   -   ^
q   w   e   r   t   y   u   i   o   p   @   [
a   s   d   f   g   h   j   k   l   ;   :   ]
z   x   c   v   b   n   m   ,   .   /   \\
f1  f2  f3  f4  f5  f6  f7  f8  f9  f10 f11 f12
num0   num1   num2   num3   num4
num5   num6   num7   num8   num9
numlock   insert   delete   home   end   pgup   pgdn
prtscr   scrolllock   pause   up   down   left   right
backspace   enter   return   capslock   nonconvert   convert
space   kana   kanji   apps   win   command   option
alt   tab   ctrl   shift   esc   fn
"""
    try:
        print(text)
        val = input('Choose targer key : ')
        spdval = input('Enter type interval : ')
        spd = decimal.Decimal(spdval)
        print('Your insert key is ... ' + val)
    except decimal.InvalidOperation:
        print('Prease enter numbers')
        main()

    num_of_secs = 5
    while num_of_secs:
        print(num_of_secs)
        time.sleep(1)
        num_of_secs -= 1

    print('start')

    while True:
        pyautogui.keyDown(val)
        pyautogui.keyUp(val)
        time.sleep(spd)

if __name__ == "__main__":
    main()
