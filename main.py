import time

from aardvark_py import *
from array import array
import pandas as pd
import sys
import os


def bar(progress, total, width=50):
    proportion = (progress + 1) / total
    filled_length = int(proportion * width)
    filled = '#' * filled_length

    blank_length = width - filled_length
    blank = ' ' * blank_length

    return f'[{filled}{blank}] {proportion * 100:.0f}%'


def read(file):
    dataframe = pd.read_excel(file, dtype={
        'Address (Hex)': str,
        'Offset (Hex)': str,
        'Data (Hex)': str,
        'Delay (ms)': float
    })
    addresses_ = dataframe['Address (Hex)'].values
    offsets_ = dataframe['Offset (Hex)'].values
    data_ = dataframe['Data (Hex)'].values
    delays_ = dataframe['Delay (ms)'].values
    return addresses_, offsets_, data_, delays_


# def configure(aardvark):
print("\nConfiguring Aardvark", end='')
aardvark = aa_open(0)
if aardvark <= 0:
    print("\nUnable to open Aardvark device on port %d" % 0)
    print("Error code = %d" % aardvark)
    input("\nPress Enter to quit.")
    sys.exit()
else:
    aa_configure(aardvark, AA_CONFIG_SPI_I2C)
    bitrate = aa_i2c_bitrate(aardvark, 100)  # Set the bitrate to 100 kHz
    bus_timeout = aa_i2c_bus_timeout(aardvark, 150)  # Set the timeout to 150 ms
    print(" - DONE\n")

file = 'AardvarkData.xlsx'
path = input(f'Hit enter to use "{file}" or enter alternative file path: ')
addresses, offsets, data, delays, errors = [], [], [], [], []
while True:
    if path != '':
        file = path.strip('" ')

    try:
        addresses, offsets, data, delays = read(file)
        print(f'Read "{file}" successfully - {len(data)} lines')
        break

    except Exception as e:
        path = input(f'"{file}" bad or not available, enter alternative file path, or hit enter to try again: ')

print()
current_address = '00'
current_delay = 0
for i, (address, offset, value, delay) in enumerate(zip(addresses, offsets, data, delays)):
    if pd.notnull(address):
        current_address = address

    if pd.notnull(delay):
        current_delay = delay

    if pd.isnull(offset) or pd.isnull(value):
        continue

    try:
        if aa_i2c_write(aardvark, int(current_address, 16), AA_I2C_NO_FLAGS,
                        array('B', [int(str(offset), 16), int(str(value), 16)])) < 1:
            raise ValueError
        else:
            print(f'\rWrote 0x{value} to 0x{offset}     {bar(i + 1, len(data))}', end='')
    except ValueError:
        print(f'\rError writing 0x{value} to 0x{offset}')
        errors.append(offset)
    time.sleep(current_delay / 1000)

if len(errors) == 0:
    print("\n\nSuccessful write without errors.")
else:
    print(f"\n\nWrite finished with {len(errors)} errors.")

# aa_i2c_write(aardvark, 0x50, AA_I2C_NO_STOP, array('B', [0x80]))
# (result, data_in) = aa_i2c_read(aardvark, 0x50, AA_I2C_NO_FLAGS, 14)
#
# print("\nRead back: ", end='')
# if result < 0:
#     print("\nFailed to read data from I2C device.")
# elif result == 0:
#     print("\nNo bytes read. Possible issue with the I2C bus.")
# else:
#     print("".join(f"{chr(data)}" for data in data_in))


aa_close(aardvark)
input("\nDone. Press Enter to quit")
