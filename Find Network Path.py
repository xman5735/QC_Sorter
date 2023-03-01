import wmi

# Create a connection to the WMI service
c = wmi.WMI()

# Query the Win32_NetworkConnection class to get the network locations
locations = c.Win32_NetworkConnection()

# Print the network locations
print('Network Locations:')
for location in locations:
    print(location.RemoteName)
