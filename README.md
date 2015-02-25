# VbsADLogonTemplate
Active Directory Logon Script Template (VBScript)

The AD logon script template maps network drives and printers based on the user's membership in so-called "map groups". Since a map group delivers the information which resource should be mapped for its members, you hardly ever need to modify this script due to changes in the drive or printer mappings.

For each drive and printer mapping you need to create an according AD group that follows an naming convention. By default, a map group for a network drive begins with "MAP-DRV-", and the prefix for a network printer map group is "MAP-PRN-". (Both prefixes can be customized by changing the constants MAP_DRIVE_GROUP_PREFIX and MAP_PRINTER_GROUP_PREFIX in this script.)

In addition to follow the naming convention you have to specify the network resources that should be mapped in the description field of a map group. In case of the map group for a network drive you must specify the drive letter followed by the unc path (seperated by a space character). In case of the printer map group you must specify the network printer's unc path.

Since the script recognized indirect or nested group memberships you are able to add users as well as groups to the map groups. The LoadGroups function contains slightly modified code that I found on Richard L. Mueller's website (<a href="http://www.rlmueller.net">http://www.rlmueller.net</a>), Thanks.
