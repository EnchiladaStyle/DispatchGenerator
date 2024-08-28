# TOUR BUS DISPATCH GENERATOR

<img src="./officialLogo.ico" width=800>

# Overview

This software solves the unique vehicle routing problem faced by tour bus companies. It takes input as tour times, tour locations, and available vehicles. It uses this input to instantly create the schedule for every vehicle and driver. It works with time window constraints, pickup/delivery constraints, and vehicle capacity constraints.

I made this software for the Cape Fox Corporation. I used to work for them as a bus driver and a tour guide. This software automates hours of scheduling, and saves the company money by providing the optimal route for each vehicle. 

[Software Demo Video](https://us06web.zoom.us/rec/share/5ZQ5HQ7kgTJZKfwJeSO9CprlXeHbokEHGetnoe6vxPtv3SGHgo4KalVQKS6Cs3Na.SrWqF1WjEwgpiwLF?startTime=1720397826000)

# Development Environment

The main UI was created with wxPython. Microsoft Excel is also used with openPyXL to handle input and output.

Google OR Tools was used for vehicle routing.

SQLite was used to store data.

Deployed with PyInstaller

# Useful Websites

- [Google OR Tools](https://developers.google.com/optimization)
