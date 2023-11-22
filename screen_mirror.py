import sys
import cv2
import time
import queue
import ctypes
import colorsys
import win32api
import win32gui
import threading
import numpy as np
from   mss import mss
from   PIL import Image
from   cuesdk import CueSdk
from   ctypes import windll
from   threading import Thread
from   numpy import uint8, uint16
from   cuesdk.structs import CorsairLedColor

# Set DPI awareness
ctypes.windll.user32.SetProcessDPIAware(1)

# Constants
shrink = 4
fps = 3
calcver = 1
lasttime = time.perf_counter()
numdone = 30
difthreshold = 3 * (100**2)
streak = 0

# Get monitor ID
def getmonitorid():
    monitors = {int(i[0]): index for index, i in enumerate(win32api.EnumDisplayMonitors())}
    winID = win32gui.GetForegroundWindow()
    MONITOR_DEFAULTTONULL = 0
    MONITOR_DEFAULTTOPRIMARY = 1
    MONITOR_DEFAULTTONEAREST = 2
    monitorID = int(win32api.MonitorFromWindow(winID, MONITOR_DEFAULTTONEAREST))
    return monitors[monitorID]

print(getmonitorid())

last2 = [np.array([255] * 4)] * 3
last1 = [None] * 3

# Functions
def tim(timed, output=True):
    global lasttime
    if not timed:
        return
    if output:
        print('{0:.20f}'.format(time.perf_counter() - lasttime))
    lasttime = time.perf_counter()

def diff(image1, image2):
    result = sum(sum(sum((image1 == image2) == False)))
    return result

def logify(value, base=0.5):
    if base <= 0 or base > 1:
        raise ValueError("The base must be between 0 and 1")
    else:
        base = base * 0.366
        return np.multiply(np.log(base * value) / np.log(base), value)

def calculate(image, index):
    timed = False
    if timed:
        print("starting")
    tim(timed, output=False)
    imagesq = image ** 2
    tim(timed)
    diff = np.abs(image[:, :, 0] - image[:, :, 1]) + \
           np.abs(image[:, :, 1] - image[:, :, 2]) + \
           np.abs(image[:, :, 0] - image[:, :, 2])
    for threshold in range(150, 10, -10):
        tim(timed)
        if timed:
            print("loop")
        mask = diff > threshold
        tim(timed)
        # Count the number of passing pixels
        pixel_count = np.sum(mask)
        tim(timed)
        if pixel_count > (15000 // (shrink ** 2)):
            break
    else:
        mean_color = np.mean(image, axis=(0, 1))
        if timed:
            print("loop over")
        tim(timed)

    if threshold > 20:
        mean_color = np.mean(image[mask], axis=(0))
        if timed:
            print("loop over")
        tim(timed)

    hsv = np.array(colorsys.rgb_to_hsv(*mean_color / 255.))
    tim(timed)
    hsv[2] = logify(hsv[2], base=0.8)
    tim(timed)
    result = np.append(np.array(np.rint(np.multiply(colorsys.hsv_to_rgb(hsv[0], hsv[1], hsv[2]), 255)), dtype=uint16), 255)
    tim(timed)
    return result

def process(image, returnlist, index):
    global last1
    global last2
    global numdone
    global streak

    if numdone < 30:
        numdone += 1
    else:
        if np.array(last1[index]).any():
            if diff(last1[index], image) < difthreshold // (shrink ** 2):
                streak += 1
                returnlist[index] = last2[index]
                return
    streak = 0
    last1[index] = image
    result = calculate(image, index)
    last2[index] = result
    returnlist[index] = result

def read_keys(input_queue):
    while True:
        input_str = input()
        input_queue.put(input_str)

def get_available_leds():
    leds = list()
    device_count = sdk.get_device_count()
    for device_index in range(device_count):
        led_positions = sdk.get_led_positions_by_device_index(device_index)
        led_colors = list(
            [CorsairLedColor(led, 0, 0, 0) for led in led_positions.keys()])
        leds.append(led_colors)
    return [sorted(i, key=lambda x: str(x).split("Id.Oem")[1].split(":")[0]) for i in leds]

def perform_pulse_effect(wave_duration, all_leds, start, end):
    global start1
    global end1
    start1 = start
    end1 = end
    time_per_frame = 50
    x = 0
    cnt = len(all_leds)
    dx = time_per_frame / wave_duration

    if bool(start == end):
        for di in range(cnt):
            device_leds = all_leds[di]
            for i, led in enumerate(device_leds):
                r, g, b, a = start[i]
                led.r = r
                led.g = g
                led.b = b
        time.sleep(wave_duration * 2)
        return

    while x < 2:
        val = x / 2
        displayed = start * (1 - val) + end * val
        displayed = [[int(i) for i in j] for j in displayed]

        for di in range(cnt):
            device_leds = all_leds[di]
            for i, led in enumerate(device_leds):
                r, g, b, a = displayed[i]
                led.r = r
                led.g = g
                led.b = b
            sdk.set_led_colors_buffer_by_device_index(di, device_leds)

        sdk.set_led_colors_flush_buffer()
        x += dx
        time.sleep(time_per_frame / 1000)

def main():
    global sdk
    global shrink

    input_queue = queue.Queue()
    input_thread = threading.Thread(target=read_keys,
                                    args=(input_queue, ),
                                    daemon=True)
    input_thread.start()

    while True:
        sdk = CueSdk()
        connected = sdk.connect()

        if not connected:
            err = sdk.get_last_error()
            print("Handshake failed: %s" % err)
            time.sleep(10)
            continue
        break

    time.sleep(0.8)
    wave_duration = 500 // fps
    colors = get_available_leds()
    print(colors)

    if not colors:
        raise Exception("Keyboard not found")

    print('Working')
    with mss() as sct:
        while True:
            totalstart = time.perf_counter()
            connected = sdk.connect()

            if not connected:
                time.sleep(10)
                sdk = CueSdk()
                continue

            if input_queue.qsize() > 0:
                input_str = input_queue.get()

                if input_str.lower() == "q":
                    print("Exiting.")
                    break
                elif input_str == "+":
                    if wave_duration > 100:
                        wave_duration -= 100
                elif input_str == "-":
                    if wave_duration < 2000:
                        wave_duration += 100

            start = time.perf_counter()
            monitor = {"top": 23, "left": 3 + 1920 * getmonitorid(), "width": 1914, "height": 1017}
            shot = sct.grab(monitor)
            shot = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
            shape = shot.size
            shot = shot.resize((shape[0] // shrink // 3 * 3, shape[1] // shrink)).convert("RGB")
            npshot = np.array(shot, dtype=uint16)
            regions = np.split(npshot, 3, axis=1)
            threads = []
            results = np.array([None] * 3)

            for index, i in enumerate(regions):
                t = Thread(target=process, args=(i, results, index))
                threads.append(t)
                t.start()

            for i, t in enumerate(threads):
                t.join()

            delay = wave_duration / 500

            if streak > 50:
                delay *= 20
            elif streak > 40:
                delay *= 10
            elif streak > 25:
                delay *= 5

            elapsed = time.perf_counter() - start
            print(elapsed)

            if elapsed < delay:
                time.sleep(delay - elapsed)

            try:
                last
            except:
                last = results
                continue

            try:
                lastwave.join()
            except Exception as err:
                print("Starting...")

            lastwave = Thread(target=perform_pulse_effect, args=(wave_duration, colors, last, results))
            lastwave.start()
            last = results

if __name__ == "__main__":
    args = list(sys.argv)
    args = [arg for arg in args if not ("py" in arg)]

    if len(args) > 0:
        time.sleep(120)

    try:
        main()
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        fullmessage = str(str(time.asctime(time.localtime(time.time()))) + ": Error %s (Message: %s) in file \"%s\" line %s" % (
        exc_type, str(e), fname, exc_tb.tb_lineno))
        print(fullmessage)
        f = open("crash_report.txt", "w")
        f.write(fullmessage)
        f.close()
