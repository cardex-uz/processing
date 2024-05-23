import os

from win32com import client as win_client


if __name__ == '__main__':
    ai = win_client.Dispatch("Illustrator.Application")
    ps = win_client.Dispatch("Photoshop.Application")

