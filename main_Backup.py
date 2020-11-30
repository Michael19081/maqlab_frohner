# Press Umschalt+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import paho.mqtt.client as mqtt
from threading import Thread
import math
import time
from pydispatch import dispatcher
from datetime import datetime


def threaded_function(_client):
    while True:
        _client.loop()


def on_connect(_client, userdata, flags, rc):
    print("Connected with result code " + str(rc))
    _client.subscribe("maqlab/#")


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

    message_splitted = message.split("|")
    t = float(message_splitted[1])
    t = t+7200  #+2h aufgrund falscher Zeitzone
    print(datetime.utcfromtimestamp(t).strftime('%Y-%m-%d %H:%M:%S')) #unix timestamp in Datum und Zeit


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    client = mqtt.Client()
    client.on_connect = on_connect
    client.on_message = on_message
    client.username_pw_set("maqlab", "maqlab")
    client.connect("techfit.at", 1883, 60)

    thread = Thread(target=threaded_function, args=(client,))
    thread.start()

    dispatcher.connect(receive_handler, signal="receive", sender="mqtt")

    

    while True:
        time.sleep(1)
        print("Hauptschleife")

    # thread.join()
    # print("thread finished...exiting")
