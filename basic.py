import time
import xlwings as xw
import xlwings.utils as xwu
import threading
import os
import paho.mqtt.client as mqtt
from threading import Thread
import math
from pydispatch import dispatcher
from datetime import datetime
import queue

q = queue.Queue()
run_once = True

py_filename_without_extension = ""
py_filename = ""
xl_filename = ""

stop_thread = False
error_cell_address = "A1"
status_cell_address = "B1"
active_devices = []
client = None
accessnr = None
wertzahl = None
wertzahl = []


# --------------------------------------------------------------------------
def main():
    global py_filename_without_extension
    global py_filename
    global xl_filename
    global client

    py_filename = os.path.basename(__file__)
    # check .py extension
    py_filename_without_extension = py_filename.split(".")[0:-1][0]
    # find the excel file in the current dir
    # we are looking for xlsx and xlsm extension
    # as result we get the first occurrence of one of this
    # excel-files, so make sure that you will not have both
    # of them in the directory
    files = os.listdir(os.path.dirname(__file__))
    for f in files:
        fn = f.split(".")[0:-1][0]
        ex = f.split(".")[-1:][0]
        if fn == py_filename_without_extension:
            if ex == 'xlsx' or ex == 'xlsm':
                xl_filename = f
                break
    print("UDF-Server started")
    print("Filepath:", __file__)
    print("Python-Filename:", py_filename)
    print("Excel-Filename:", xl_filename)
    '''
    source.range('A1').expand().clear_contents()
    source.range('A1').value = cursor.fetchall()
    '''
    wb = xw.Book.caller()
    wb.sheets.active.api.OLEObjects("MessageBox").Object.Value = "Initialized sucessfully\n"

    client = mqtt.Client()
    client.on_connect = on_connect
    client.on_message = on_message
    client.username_pw_set("maqlab", "maqlab")
    client.connect("techfit.at", 1883, 60)

    thread = Thread(target=mqtt_loop, args=(client,))
    thread.start()

    dispatcher.connect(receive_handler, signal="receive", sender="mqtt")
    time.sleep(1)
    active_devices.clear()
    client.publish("maqlab/user1/cmd/?")
    time.sleep(1)
    print(active_devices)


# --------------------------------------------------------------------------
def start(interval, count):
    if xl_filename == (""):
        main()

    wb = xw.Book.caller()
    wb.sheets.active.range(error_cell_address).value = ""
    wb.sheets.active.api.OLEObjects("MessageBox").Object.Value = "Started...\n"

    combo = 'ComboBox1'
    wb.sheets.active.api.OLEObjects(combo).Object.AddItem('NTP-5431')  # active = aktuelle Tabelle
    wb.sheets.active.api.OLEObjects(combo).Object.AddItem("SM2400")
    wb.sheets.active.api.OLEObjects(combo).Object.AddItem("BK-E2831")
    # wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnCount = 2
    # wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnWidths = 0

    # wb.sheets.active.api.OLEObjects("ComboBox").Object.Value = "some value"
    if isinstance(interval, str):
        try:
            interval = interval.replace(",", ".")  # Komma von , auf . ändern
            interval = float(interval)
        except:
            interval = 0
            wb.sheets.active.range(error_cell_address).value = "ERROR: Interval - invalid value format"
            return

    if isinstance(count, str):
        try:
            count = int(count)
            if count <= 0:
                wb.sheets.active.range(error_cell_address).value = "WARNING: Number of measure cycles is zero"
        except:
            wb.sheets.active.range(error_cell_address).value = "ERROR: Count - invalid value format"
            return

    # starting the measuring thread
    global stop_thread
    stop_thread = False
    x = threading.Thread(target=measure, args=(float(interval), int(count),))
    x.start()


# --------------------------------------------------------------------------
def stop():
    global stop_thread
    stop_thread = True


# --------------------------------------------------------------------------
def measure(t, count):
    global stop_thread
    global client
    global accessnr
    global wertzahl

    # accessing the sheet with index 0
    sht = xw.Book(xl_filename).sheets[0]  # sheet 0 = 1.Tabelle

    # Status zelle Text ändern und umfärben
    sht.range(status_cell_address).api.Font.Color = xwu.rgb_to_int((0, 0, 0))  # font color of text
    sht.range(status_cell_address).color = xwu.rgb_to_int((0, 200, 10))  # cell color
    sht.range(status_cell_address).value = "Messung läuft"  # cell text
    print("Thread Running...")
    i = 0
    while i < count:
        print(i)
        sht.range("J5").value = str(i)
        sht.range("J6").value = str(count)

        # Index der Messung in Spalte A (=1)
        # das ist eine weitere Möglichkeit eine Zelle zu adressieren
        # ( Zeilennummer, Spaltennummer)
        cell = (i + 3, 1)
        sht.range(cell).value = str(i)

        # vorherige Zelle färben
        if i > 0:
            sht.range(i + 2, 1).color = xwu.rgb_to_int((100, 100, 100))

        # actuelle Zelle färben
        sht.range(cell).color = xwu.rgb_to_int((0, 200, 10))

        # -------------------------------------------

        global run_once
        if run_once:
            run_once = False
            Voltage = 0
            Current = 0
            Zeile = 3

        # sht.range('E32').value = [['TabKopfX', 'TabKopfY'], [1, 2], [10, 20]]

        sht.range('C22').value = 'Verfuegbar'
        sht.range('C23').value = active_devices
        accessnr = sht.range('L15').value
        accessnr = int(accessnr)

        sp1 = (sht["N14"].value)
        wert1 = "maqlab/user1/cmd/" + str(accessnr) + "/" + str(sp1) + "?"
        #print(wert1)
        client.publish(topic=wert1, payload="");

        cell = (i + 16, 14)
        sht.range(cell).value = wertzahl

        '''
        sp2 = (sht["O14"].value)

        '''

        # sht.range('N15').value = sp1_wert

        # -------------------------------------------

        # timing
        timer = 0
        while timer < t:
            if stop_thread:
                break
            time.sleep(0.01)
            timer = timer + 0.01
        if stop_thread:
            stop_thread = False
            break
        i += 1  # nächste Zeile

    # Statuszelle Text ändern und umfärben
    sht.range(status_cell_address).api.Font.Color = xwu.rgb_to_int((255, 255, 255))
    sht.range(status_cell_address).color = xwu.rgb_to_int((200, 10, 10))
    sht.range(status_cell_address).value = "Messung gestopped"
    print("Thread Stopped")


# --------------------------------------------------------------------------
# main()

def mqtt_loop(_client):
    while True:
        _client.loop()


def on_connect(_client, userdata, flags, rc):
    global active_devices
    print("Connected with result code " + str(rc))
    _client.subscribe("maqlab/#")

    # _client.subscribe("maqlab/+/rep/#")


def on_message(_client, userdata, msg):
    # print(msg.topic + " " + str(msg.payload))

    dispatcher.send(message=str(msg.topic) + "|" + msg.payload.decode("utf-8"),
                    signal="receive",
                    sender="mqtt")


# ----------------------------------------------------------------
# D I S P A T C H - H A N D L E R
# ----------------------------------------------------------------

def receive_handler(message):
    # print('message in receive_handler: {}'.format(message))
    global run_once
    global starttime
    global active_devices
    global accessnr
    global wertzahl
    # print(message)
    if message.split("|")[0] == 'maqlab/ping/':
        message_splitted = message.split("|")
        t = float(message_splitted[1])
        t = t + 7200  # +2h aufgrund falscher Zeitzone

        if run_once:  # globale Variable, dass man nur einmal in die Funktion kommt
            run_once = False
            starttime = t
            print(datetime.utcfromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S'))  # unix timestamp in Datum und Zeit

        timex = t - starttime
        #print(int(timex), "s")

    # print(str(message.split("|")[0]))

    if '/accessnumber' in message:
        accessnumber = message.split("|")[1]
        topic = message.split("|")[0]
        device_name = topic.split("/")[3]
        # message_new = message.replace('maqlab/user1/rep/', '')
        # message_new = message_new.replace('/accessnumber', '')
        # print(device_name, accessnumber)
        active_devices.append((device_name, accessnumber))
        # exec("%s = %d" % (message_new.split("|")[0], int(message_new.split("|")[1])))     # Topic als Variable ihren Wert zuweisen
        # print(message_new)

    if str("/rep/" + str(accessnr)) in message:      #Werte von gewählter Größe abfragen
        wertzahl = message.split("|")[1]
        # wertzahl.append(message.split("|")[1])
        #werttopic = message.split("/")[3]          #bereits in wertzahl enthalten


    # vars()[message_new.split("|")[0]] = message_new.split("|")[1]
    # print(message_new.split("|")[0])


# Press the green button in the gutter to run the script.


if __name__ == '__main__':
    client = mqtt.Client()
    client.on_connect = on_connect
    client.on_message = on_message
    client.username_pw_set("maqlab", "maqlab")
    client.connect("techfit.at", 1883, 60)

    thread = Thread(target=mqtt_loop, args=(client,))
    thread.start()

    dispatcher.connect(receive_handler, signal="receive", sender="mqtt")
    time.sleep(1)
    active_devices.clear()
    client.publish("maqlab/user1/cmd/?")
    time.sleep(1)
    print(active_devices)
    while True:
        time.sleep(1)
        # client.publish("topic", "payload");
        #print("Hauptschleife")

    # thread.join()
    # print("thread finished...exiting")
