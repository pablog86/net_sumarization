import ipaddress

import tkinter as tk
from tkinter import filedialog
import os
from os import listdir
from os.path import isfile, join

import openpyxl

def sumarizar_subredes(subredes):
    ordered_subnets = sorted(subredes)
    print("Input:")
    print(ordered_subnets)
    current_subnet = ordered_subnets[0]
    flag = 1

    while flag == 1:
    	summarized = set()
    	current_subnet = ordered_subnets[0]
    	flag = 0
    	print("--------- Pass ---------")
    	for subred in ordered_subnets[1:]:
    	    if subred.network_address == current_subnet.broadcast_address + 1 and subred.prefixlen == current_subnet.prefixlen:
    	        print("Contiguous networks: ", str(current_subnet), " = ", str(subred))
    	        current_subnet = subred.supernet()
    	        print("Supernet: ", str(current_subnet))
    	        flag = 1
    	    else:
    	        summarized.add(current_subnet)
    	        current_subnet = subred
    	print("Qty of networks: ", len(summarized))
    	print("Pass sumarization: ", summarized)
    	print("")
    	ordered_subnets = list(summarized)
    	ordered_subnets = sorted(ordered_subnets)

    summarized.add(current_subnet)	#Last member
    summarized = [str(address) for address in summarized]
    return list(summarized)



if __name__ == "__main__":

	root = tk.Tk()
	root.withdraw()
	path = filedialog.askopenfilename()
	wb = openpyxl.load_workbook(path)
	sheet = wb.active
	max_row = sheet.max_row

	subnets = []
	subnets = [c.value for c in sheet['A']]
	subnets_ip = [ipaddress.IPv4Network(subnet) for subnet in subnets]
	summary = sumarizar_subredes(subnets_ip)
	print("Total length of the input: ", len(subnets_ip))
	print("Summary length: ", len(summary))
	print("Output: ") 
	print(summary)
	print("List: ")

	for indice, net in enumerate(summary):
		print(net)
		celda = sheet.cell(row=indice+1, column=3)
		celda.value = net
	wb.save(path)
