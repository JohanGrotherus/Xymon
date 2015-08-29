#                                 #
# Xymon addons in this directory  #
#                                 #
################################### 

1. replications.vbs

This script prints basic domain information and then uses WMI queries to check for errors in replication from
the particular domain controller. 

To use the script, add the script in the BBWin external folder and add the following to the externals section
of bbwin.cfg

< load value="..\ext\replication.vbs" timer="300s" />
