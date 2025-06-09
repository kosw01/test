# -*- coding: utf-8 -*-
"""
Created on Mon Jan  8 15:26:07 2024

@author: Y15599
"""
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
filepath = 'collect_data.csv'
savepath = 'collect_data_fft.csv'
df = pd.read_csv(filepath, encoding='ISO-8859-1')



n = len(df)
time = np.arange(0, n/100, 0.01)

k = np.arange(n)
Fs = 1 / 0.01
T = n / Fs
freq = k / T
freq = freq[range(int(n / 2))]

Y = []
Y_abs = []

# Iterate over column indices
for i in range(len(df.columns)):
    Y.append(np.fft.fft(df.iloc[:, i]) / n)
    Y[i] = Y[i][range(int(n / 2))]
    Y_abs.append(abs(Y[i]))

    fig, ax = plt.subplots(2, 1, figsize=(12, 8))
    ax[0].plot(time, df.iloc[:, i])
    ax[0].set_title(f'{df.columns[i]}')
    ax[0].set_xlabel('Time')
    ax[0].set_ylabel('Amplitude')
    ax[1].set_xlim([0,60])
    ax[0].grid(True)
    ax[0].set_xlim([0, 40])
    
    ax[1].grid(True)
    ax[1].plot(freq[10:], abs(Y[i][10:]), 'r')
    ax[1].set_xlabel('Freq (Hz)')
    ax[1].set_ylabel('|Y(freq)|')
    ax[1].set_xlim([0,5])
    #ax[1].set_ylim([0, 0.03])
    ax[1].grid(True)
    plt.savefig(f'{df.columns[i]}.png')
    #plt.show()

# Save Y_abs as a single CSV file
Y_abs_combined = pd.DataFrame(Y_abs).T
Y_abs_combined.columns = df.columns
Y_abs_combined.to_csv(savepath, index= False)
