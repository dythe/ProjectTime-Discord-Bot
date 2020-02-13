# Work with Python 3.6
import re
import discord
import requests
import gspread
import time
import asyncio
import xlsxwriter
import json
import sys
import os
from collections import OrderedDict, defaultdict
from authlib.client import AssertionSession
from gspread import Client

client = discord.Client()

# Google Sheets and Discord authentication
with open("settings.json", "r") as jsonFile:
    data = json.load(jsonFile)

discordToken = data['discordToken']
googleAPIKey = data['googleAPIKey']

# Your Google Keyfile Configuration location
keyFileLocation = data['keyFileLocation']

# Channel and Server ID
serverID = data['serverID']
server = client.get_server(id=serverID)

tyrfasID = data['tyrfasID']
velkazarID = data['velkazarID']
lakreilID = data['lakreilID']
channel = data['channel']
botCommandsID = data['botCommandsID']

authenticationFlag = 0
authenticationFailCount = 0

# Members
# PTMemberArraywID = []
PTMemberArray = []
PTTrialArray = []
PTModArray = []

# Groups based on row number in Google Sheets

#Tyrfas
group1Rows = [11, 12, 13]

#Lakreil
group2Rows = [22, 23, 24]
group3Rows = [25, 26, 27]

#Velkazar
group4Rows = [36, 37, 38]
group5Rows = [39, 40, 41]
group6Rows = [42, 43, 44]
group7Rows = [45, 46, 47]
group8Rows = [48, 49, 50]
group9Rows = [51, 52, 53]
group10Rows = [54, 55, 56]
group11Rows = [57, 58, 59]
group12Rows = [60, 61, 62]
group13Rows = [63, 64, 65]
group14Rows = [66, 67, 68]
group15Rows = [69, 70, 71]
group16Rows = [72, 73, 74]
group17Rows = [75, 76, 77]
group18Rows = [78, 79, 80]
group19Rows = [81, 82, 83]
group20Rows = [84, 85, 86]
group21Rows = [87, 88, 89]
group22Rows = [90, 91, 92]
group23Rows = [93, 94, 95]
group24Rows = [96, 97, 98]
group25Rows = [99, 100, 101]

groupsDict = [
    (1, group1Rows),
    (2, group2Rows),
    (3, group3Rows),
    (4, group4Rows),
    (5, group5Rows),
    (6, group6Rows),
    (7, group7Rows),
    (8, group8Rows),
    (9, group9Rows),
    (10, group10Rows),
    (11, group11Rows),
    (12, group12Rows),
    (13, group13Rows),
    (14, group14Rows),
    (15, group15Rows),
    (16, group16Rows),
    (17, group17Rows),
    (18, group18Rows),
    (19, group19Rows),
    (20, group20Rows),
    (21, group21Rows),
    (22, group22Rows),
    (23, group23Rows),
    (24, group24Rows),
    (25, group25Rows),

]

# Group Numbers
tyrfasGroups = [1]
lakreilGroups = [2, 3]
velkazarGroups = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]

# Naming
tyrfasNameTag = 'Guild Conquest 1 (Tyrfas)'
tyrfasName = 'Tyrfas'

lakreilNameTag = 'Guild Conquest 2 (Lakreil)'
lakreilName = 'Lakreil'

velkazarNameTag = 'Guild Conquest 3 (Velkazar)'
velkazarName = 'Velkazar'

currentGC = None
bossName = None
# lookup_list = []
scope = None
credentials = None
gc = None
sht1 = None
guildConquestSheet = None
guildWarsSheet = None
worksheetSettings = None
discordID = None

@client.event
async def on_message(message):

    # global scope
    # global credentials
    # global gc
    # global sht1
    # global guildConquestSheet
    # global worksheetSettings
    # global discordID
    global channel

    if message.author == client.user:
        return

    url = None
    comment = None
    user = message.author
    screenshotFlag = 0
    commentsFlag = 0

    if not message.content.startswith('!'):
        # print('invalid')
        pass

    allowedChannels = ['bot-commands', 'pt-general', 'guild-conquest-gc2', 'guild-conquest-gc3']
    # print (currevents)
    # print (message.channel)

    if str(message.channel) == 'bot-testing-channel' and message.content.startswith('!sayd') and (user.id in PTMemberArray or user.id in PTTrialArray):
        # await client.delete_message(message)
        
        try:
            channelToSend = message.content.split(" ", 2)[1]
        except:
            return

        try:
            idToMention = '<@' + message.content.split(" ")[2] + '>'
        except:
            return

        try:
            msgToSend = message.content.split(" ", 3)[3]
        except:
            msgToSend = ""

        await client.send_message(discord.Object(id=channelToSend), "%s %s" % (idToMention, msgToSend))
        print("%s pinged" % (message.author))
        return
    if not message.channel.is_private and message.content.startswith('!sayd') and (user.id in PTMemberArray or user.id in PTTrialArray):
        await client.delete_message(message)
        idToMention = '<@' + message.content.split(" ")[1] + '>'

        try:
            msgToSend = message.content.split(" ", 2)[2]
        except:
            msgToSend = ""

        await client.send_message(message.channel, "%s %s" % (idToMention, msgToSend))
        print("%s pinged" % (message.author))
        return
    elif message.content.startswith('!damage') and (user.id in PTMemberArray or user.id in PTTrialArray) and str(message.channel) in allowedChannels:
        try:
            gcType = message.content.split(" ")[1]
        except:
            gcType = ""
        if gcType == "":
            await client.send_message(message.channel, "Retrieving current Guild Conquest damage status...")
            asyncio.run_coroutine_threadsafe(retrieveDmg(user, message, message.channel), client.loop)
        else:

            if (user.id in PTModArray or user.id == '141672556125093889'):
                try:
                    damageNum = message.content.split(" ")[2]
                except:
                    damageNum = ""

                if damageNum == "":
                    await client.send_message(message.channel, "No damage value specified!")
                else:
                    with open("settings.json", "r") as jsonFile:
                        data = json.load(jsonFile)

                    if gcType == "gc1":
                        data['gc1HistoricalHighest'] = damageNum
                        await client.send_message(message.channel, "Tyrfas Historical Damage Score is set to %s" % damageNum)
                    elif gcType == "gc2":
                        data['gc2HistoricalHighest'] = damageNum
                        await client.send_message(message.channel, "Lakreil Historical Damage Score is set to %s" % damageNum)
                    elif gcType == "gc3":
                        data['gc3HistoricalHighest'] = damageNum
                        await client.send_message(message.channel, "Velkazar Historical Damage Score is set to %s" % damageNum)

                with open("settings.json", "w") as jsonFile:
                    json.dump(data, jsonFile, indent=4)

    elif message.content == '!group' and (user.id in PTMemberArray or user.id in PTTrialArray) and str(message.channel) in allowedChannels:
        if not message.channel.is_private:
            await client.send_message(message.channel, message.author.mention + " Your team composition will be sent to you via PM.")
            
        await client.send_message(message.author, "Searching for your team compositions, Please wait...")
        print(str(message.author) + " requested for their groupings.")
        asyncio.run_coroutine_threadsafe(retrieveTeamComp(user, message, 0, 0), client.loop)
        return
    elif message.content.startswith('!retrieve') and (user.id in PTModArray or user.id == '141672556125093889'):
        currentUserID = '<@' + message.mentions[0].id + '>'
        await client.send_message(message.author, "Retrieving team compositions for %s" % (currentUserID))
        await client.send_message(message.channel, "The team composition of %s will be sent to you via PM." % (currentUserID))
        print("%s requested for %s groups." % (message.author, message.mentions[0].name))
        asyncio.run_coroutine_threadsafe(retrieveTeamComp(user, message, 1, message.mentions[0].id), client.loop)

        # print (message.mentions[0].id)
    elif message.content.startswith('!restart') and (user.id == '123631705662685188'):
        await client.send_message(message.channel, 'Bot is restarting.')
        os.execv(sys.executable, ['python'] + sys.argv)
    if message.content.startswith('!overview') and user.id in PTModArray:
        asyncio.run_coroutine_threadsafe(retrieveDmg(user, message), client.loop)
    elif message.content.startswith('!spreadsheet') and user.id in PTModArray:
        asyncio.run_coroutine_threadsafe(spreadsheetToggle(user, message), client.loop)
    elif message.content.startswith('!ptcommands') and (user.id in PTMemberArray or user.id in PTTrialArray):
        
        ptcommandsEmbed = discord.Embed(title="ProjectTime Bot Commands", color=0xffffff)
        ptcommandsEmbed.add_field(name='!group', value="Find out your group's team composition via PM (PT Members / PT Trial command)", inline=False)
        ptcommandsEmbed.add_field(name='!damage', value="Retrieve Current Guild Conquest damage status (PT Members / PT Trial command)", inline=False)
        ptcommandsEmbed.add_field(name='!damage (gc1/gc2/gc3) (damage no)', value="Update the historical highest score for a certain GC (Mod command)", inline=False)
        ptcommandsEmbed.add_field(name='!retrieve', value="Find out the team composition of a specific member via PM (Mod command)", inline=False)
        ptcommandsEmbed.add_field(name='!spreadsheeton/!spreadsheetoff/!spreadsheet', value="Turn on spreadsheet, Turn off spreadsheet, View spreadsheet status (Mod command)", inline=False)
        # ptcommandsEmbed.add_field(name='!gc', value="Find out current GC score channel (Mod command)", inline=False)
        # ptcommandsEmbed.add_field(name='!gc1', value="Set GC score channel to Tyrfas (Mod command)", inline=False)
        # ptcommandsEmbed.add_field(name='!gc2', value="Set GC score channel to Lakreil and do an announcement (Mod command)", inline=False)
        # ptcommandsEmbed.add_field(name='!gc3', value="Set GC score channel to Velkazar and do an announcement (Mod command)", inline=False)
        ptcommandsEmbed.add_field(name='!restart', value="Force the bot to restart (Mod command)", inline=False)
        await client.send_message(message.channel, embed=ptcommandsEmbed)
        
    # elif message.content.startswith('!gc') and (user.id in PTModArray or user.id == '141672556125093889'):

    #     data = None

    #     if message.content == '!gc':
    #         with open('settings.json') as f:
    #             data = json.load(f)
            
    #         await client.send_message(message.channel, "Current Guild Conquest scores channel is set to " + str(data["currentBoss"]))
    #     else:
       #          currentGC = message.content.split("gc")[1]
                
       #          if currentGC == '1':
       #              bossName = tyrfasName
       #              channel = tyrfasID
       #          elif currentGC == '2':
       #              bossName = lakreilName
       #              channel = lakreilID
       #          elif currentGC == '3':
       #              bossName = velkazarName
       #              channel = velkazarID
       #          else:
       #            return

       #          with open("settings.json", "r") as jsonFile:
       #            data = json.load(jsonFile)
                
       #          data['currentBoss'] = bossName

       #          with open("settings.json", "w") as jsonFile:
       #            json.dump(data, jsonFile, indent=4)
       #          # settingsFormat = {"currentBoss": bossName}
       #          # with open('settings.json', 'w') as outfile:
       #          #     x['nickname'] = newname
       #          #     json.dump(settingsFormat, outfile, indent=4)

       #          if bossName == tyrfasName:
       #              print(str(message.author) + " has set Guild Conquest scores channel to " + bossName)
       #              await client.send_message(message.channel, "Guild Conquest scores channel is set to " + bossName)
       #          else:
       #              PTMemberRoleID = '<@&458728830992121868>'
       #              PTTrialRoleID = '<@&469377991638646785>'

       #              if bossName == lakreilName:
       #                  formattedAnnouncement = "%s %s %s is down, hit %s asap..." % (PTMemberRoleID, PTTrialRoleID, tyrfasName, lakreilName)
       #              elif bossName == velkazarName:
       #                  formattedAnnouncement = "%s %s %s is down, hit %s asap..." % (PTMemberRoleID, PTTrialRoleID, lakreilName, velkazarName)

       #              print(str(message.author) + " has set Guild Conquest scores channel to " + bossName)
       #              await client.send_message(message.channel, "Guild Conquest scores channel is set to " + bossName)
       #              await client.send_message(discord.Object(id='458781346081406996'), formattedAnnouncement)


    # with open('settings.json') as f:
    #     data = json.load(f)

    # if str(data["currentBoss"]) == tyrfasName:
    #     channel = tyrfasID
    # elif str(data["currentBoss"]) == velkazarName:
    #     channel = velkazarID
    # elif str(data["currentBoss"]) == lakreilName:
    #     channel = lakreilID

    embed = None
    if message.channel.is_private and bool(re.search('!group', message.content, re.IGNORECASE)) == False:
        if message.channel.is_private and message.attachments and (user.id in PTMemberArray or user.id in PTTrialArray):
            screenshotFlag = 1
            for attach in message.attachments:
                url = attach['url']

            print(str(user) + " has uploaded a screenshot. (" + url + ")")

            # msg = 'Screenshot uploaded by ' + str(message.author)
            msg = 'Screenshot uploaded by ' + str(message.author)
            embed = discord.Embed(title=msg, url=url, color=0xf20b0b)
            embed.set_image(url=url)
            # await client.send_message(discord.Object(id=channel), msg)
            await client.send_message(discord.Object(id=channel), embed=embed)
            msg2 = 'Screenshot uploaded by ' + '<@' + message.author.id + '>'

        if message.channel.is_private and message.content and (user.id in PTMemberArray or user.id in PTTrialArray):
            commentsFlag = 1
            print(str(user) + " has commented. (" + str(message.content) + ")")

            # msg = 'Comment by ' + str(message.author) + ': ' + str(message.content)
            msg = 'Comment by ' + '<@' + message.author.id + '>' + ': ' + str(message.content)
            await client.send_message(discord.Object(id=channel), msg)            

        if screenshotFlag == 1 and commentsFlag == 0:
            await client.send_message(message.author, 'Your screenshot has been uploaded to the scores channel.')   
            # await client.send_message(message.author, '(URL: ' + str(url) + ')')
            await client.send_message(message.author, embed=embed)
        elif commentsFlag == 1 and screenshotFlag == 0:
            await client.send_message(message.author, 'Your comments has been uploaded to the scores channel.')
            await client.send_message(message.author, 'Comments: ' + str(message.content))
        elif commentsFlag == 1 and screenshotFlag == 1:
            await client.send_message(message.author, 'Your screenshot and comments has been uploaded to the scores channel.')
            # await client.send_message(message.author, '(URL: ' + str(url) + ')')
            await client.send_message(message.author, embed=embed)
            await client.send_message(message.author, 'Comments: ' + str(message.content))
    elif message.channel.is_private and bool(re.search('!group', message.content, re.IGNORECASE)) == True:
        await client.send_message(message.author, 'Please do not type the !group command in private message but do it in the main channels')


# @client.event
# async def on_reaction_add(reaction, user):
#     rolesAssignID = '504394602829053962'
#     server = client.get_server(id=serverID)

#     GRTyrfas = get(server.roles, name='GR Tyrfas')
#     GRLakreil = get(server.roles, name='GR Lakreil')

#     if reaction.message.channel.id != rolesAssignID:
#         return #So it only happens in the specified channel
#     if str(reaction.emoji) == "<:one:504396034592473133>":
#         await client.add_roles(user, GRTyrfas)
#         print ('%s has been assigned %s role' % (user, GRTyrfas))
#     elif str(reaction.emoji) == "<:two:504396184127930368>":
#         await client.add_roles(user, GRLakreil)
#         print ('%s has been assigned %s role' % (user, GRLakreil))

async def spreadsheetToggle(user, message):
    # for i in range(0,50):
    #     while True:
    #         try:
    #             scope = ['https://spreadsheets.google.com/feeds',
    #                      'https://www.googleapis.com/auth/drive']

    #             credentials = ServiceAccountCredentials.from_json_keyfile_name('PTBot.json', scope)

    #             gc = gspread.authorize(credentials)

    #             sht1 = gc.open_by_key(googleAPIKey)
                  
    #         except:
    #             print('Exception detected. retrying...')
    #             continue
    #         break
    
    # guildConquestSheet = sht1.get_worksheet(0)
    # worksheetSettings = sht1.get_worksheet(1)
    # discordID = sht1.get_worksheet(2)  

    if message.content == '!spreadsheet':
        botSwitch = guildConquestSheet.cell(2, 7).value
        if botSwitch == 'YES':
            botSwitch = 'ON'
        elif botSwitch == 'NO':
            botSwitch = 'OFF'
        tempString = 'Spreadsheet status: %s' % (botSwitch)
        await client.send_message(message.channel, tempString)
    else:
        status = message.content.split("spreadsheet")[1]

        if status.lower() == 'on':
            guildConquestSheet.update_cell(2, 7, 'YES')
            await client.send_message(message.channel, 'Spreadsheet is now online')
        elif status.lower() == 'off':
            guildConquestSheet.update_cell(2, 7, 'NO')
            await client.send_message(message.channel, 'Spreadsheet is now offline')
        else:
            await client.send_message(message.channel, 'Invalid command.')

def retrieveRows(num):

    print ('Group No: ' + str(num))
    temp = None

    computed = int(num) - 1
    temp = groupsDict[computed][1]

    print("Group Cells: " + str(temp))

    return temp

# Authenticate using Autolib AssertionSession over oauth2client
# https://blog.authlib.org/2018/authlib-for-gspread
def authWithAuthLib():
    global scope
    global credentials
    global gc
    global sht1
    global guildConquestSheet
    global worksheetSettings
    global discordID
    global guildWarsSheet

    try:
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive',
        ]

        # Keyfile from Google Configuration
        session = create_assertion_session(keyFileLocation, scope)

        # print (session)
        gc = Client(None, session)
        sht1 = gc.open_by_key(googleAPIKey)
        # guildConquestSheet = sht1.get_worksheet(0)
        # worksheetSettings = sht1.get_worksheet(1)
        # discordID = sht1.get_worksheet(2)
        
        # print (sht1.worksheets())

        guildConquestSheet = sht1.worksheet("Guild Conquest")
        worksheetSettings = sht1.worksheet("Settings")
        discordID = sht1.worksheet("Discord ID")
        guildWarsSheet = sht1.worksheet("Guild Wars")
        return 1
    except:
        return 0
    
async def retrieveTeamComp(user, message, commandType, idDiscord):

    botSwitch = guildConquestSheet.cell(2, 7).value
    print ('Bot switch status: %s' % (botSwitch))

    if botSwitch == 'NO' and commandType == 0:
        await client.send_message(message.channel, "Team composition page is under review, please try again some other time.")
        await client.send_message(message.author, "Team composition page is under review, please try again some other time.")
    elif botSwitch == 'YES' or commandType == 1:
        # Find a cell with exact string value

        if commandType == 0:
            cell = discordID.find(user.id)
        elif commandType == 1:
            cell = discordID.find(idDiscord)
        # print ('1:' + str(cell))

        name = discordID.cell(cell.row, 1).value

        # print ('2:' + str(name))
        lookup = worksheetSettings.findall(name)

        # print ('3:' + str(lookup))
        
        lookup_list = []

        for x in lookup:
            lookup_list.append(str(x.row))

        lookup_list = list(OrderedDict.fromkeys(lookup_list))
        print ('lookup_list: ' + str(lookup_list))

        # Bans embed
        GC1Bans = guildConquestSheet.cell(8, 4).value
        GC2Bans = guildConquestSheet.cell(19, 4).value
        GC3Bans = guildConquestSheet.cell(33, 4).value

        GC1Val = GC1Bans.split("Bans:")[1].strip()
        GC2Val = GC2Bans.split("Bans:")[1].strip()
        GC3Val = GC3Bans.split("Bans:")[1].strip()

        if GC1Val == '':
            GC1Val = 'No bans set yet'
        
        if GC2Val == '':
            GC2Val = 'No bans set yet'
        
        if GC3Val == '':
            GC3Val = 'No bans set yet'

        # print (GC1Val)
        # print (GC2Val)
        # print (GC3Val)

        # Only with the !group command we show them the banlist
        if commandType == 0:
            bansEmbed = discord.Embed(title="Bans", description="This is the banlist of the upcoming/current Guild Conquest", color=0xffffff)
            bansEmbed.add_field(name=tyrfasNameTag, value=GC1Val, inline=False)
            bansEmbed.add_field(name=lakreilNameTag, value=GC2Val, inline=False)
            bansEmbed.add_field(name=velkazarNameTag, value=GC3Val, inline=False)
            await client.send_message(message.author, embed=bansEmbed)

        for x in lookup_list:
            groupNo = None
            groupString = None
            rowsArray = []

            groupNo = x
            groupString = '==== Group %s ==== \n' % (groupNo)
            rowsArray = retrieveRows(x)
            
            nameRange = 'E%s:E%s' % (min(rowsArray), max(rowsArray))
            actualRange = 'F%s:F%s' % (min(rowsArray), max(rowsArray))
            heroRange = 'G%s:G%s' % (min(rowsArray), max(rowsArray))
            nameCells = guildConquestSheet.range(nameRange)
            actualCells = guildConquestSheet.range(actualRange)
            heroCells = guildConquestSheet.range(heroRange)
            timezoneCell = guildConquestSheet.cell(min(rowsArray), 13).value
            publicCommentsCell = guildConquestSheet.cell(min(rowsArray), 12).value

            if int(groupNo) in tyrfasGroups:
                assignedBoss = tyrfasNameTag
            elif int(groupNo) in velkazarGroups:
                assignedBoss = velkazarNameTag
            elif int(groupNo) in lakreilGroups:
                assignedBoss = lakreilNameTag

            descriptionFormat = "Assigned Boss: **%s** \n" % (assignedBoss)
            # descriptionFormat = descriptionFormat + "Total number of groups on %s: **%s** \n" % (assignedBoss, len(lookup_list))

            descriptionFormat = descriptionFormat + "Timezone: **%s** \n" % (timezoneCell)

            if publicCommentsCell != '':
                descriptionFormat = descriptionFormat + "Comments: \n **%s**" % (publicCommentsCell)

            embed = discord.Embed(title="Group %s" % (groupNo), description=descriptionFormat, color=0x7289da)
            # whatisURL = 'C:/Users/Jefferson/Desktop/DiscordBot/gc3.png'
            # embed.set_thumbnail(url=whatisURL)
            
            for i, x in enumerate(nameCells):
                if nameCells[i].value != '' and heroCells[i].value != '':
                    if nameCells[i].value == actualCells[i].value:
                        embed.add_field(name=nameCells[i].value, value=heroCells[i].value, inline=False)
                        groupString = groupString + nameCells[i].value + ": " + heroCells[i].value + '\n'
                    else:
                        embed.add_field(name=nameCells[i].value + " (" + actualCells[i].value + ")", value=heroCells[i].value, inline=False)
                        groupString = groupString + nameCells[i].value + " (" + actualCells[i].value + "): " + heroCells[i].value + '\n'

            await client.send_message(message.author, embed=embed)
            print ("\n" + "==========================")
            print (descriptionFormat)
            print (groupString)

            # lookup_list.remove(groupNo)

        if len(lookup_list) == 0:
            await client.send_message(message.author, "No groups has been found for you.")
            print (str(message.author) + " has no groups, check him out?")

async def retrieveDmg(user, message, currentChannel):

    # print ("%s requested for current GC damage status." % (user))
    # for i in range(0,50):
    #     while True:
    #         try:
    #             scope = ['https://spreadsheets.google.com/feeds',
    #                      'https://www.googleapis.com/auth/drive']

    #             credentials = ServiceAccountCredentials.from_json_keyfile_name('PTBot.json', scope)

    #             gc = gspread.authorize(credentials)
    #         except:
    #             print('Exception detected. retrying...')
    #             continue
    #         break
    # sht1 = gc.open_by_key(googleAPIKey)
    # guildConquestSheet = sht1.get_worksheet(0)

    gc1Lowest = guildConquestSheet.cell(5, 13).value
    gc1Highest = guildConquestSheet.cell(6, 13).value
    gc1RunsCompleted = guildConquestSheet.cell(6, 11).value
    gc1TotalRuns = guildConquestSheet.cell(5, 11).value
    gc1LowestCombined = guildConquestSheet.cell(7, 13).value
    gc1HighestCombined = guildConquestSheet.cell(8, 13).value
    gc1TotalDamageDone = guildConquestSheet.cell(8, 11).value

    gc2Lowest = guildConquestSheet.cell(22, 13).value
    gc2Highest = guildConquestSheet.cell(23, 13).value
    gc2RunsCompleted = guildConquestSheet.cell(23, 11).value
    gc2TotalRuns = guildConquestSheet.cell(22, 11).value
    gc2LowestCombined = guildConquestSheet.cell(24, 13).value
    gc2HighestCombined = guildConquestSheet.cell(25, 13).value
    gc2TotalDamageDone = guildConquestSheet.cell(25, 11).value

    gc3Lowest = guildConquestSheet.cell(81, 13).value
    gc3Highest = guildConquestSheet.cell(82, 13).value
    gc3RunsCompleted = guildConquestSheet.cell(82, 11).value
    gc3TotalRuns = guildConquestSheet.cell(81, 11).value
    gc3LowestCombined = guildConquestSheet.cell(83, 13).value
    gc3HighestCombined = guildConquestSheet.cell(84, 13).value
    gc3TotalDamageDone = guildConquestSheet.cell(84, 11).value

    with open('settings.json') as f:
        data = json.load(f)

    gc1HistoricalHighest = data["gc1HistoricalHighest"]
    gc2HistoricalHighest = data["gc2HistoricalHighest"]
    gc3HistoricalHighest = data["gc3HistoricalHighest"]
    
    gc1String = "Total Runs Completed: **%s** / **%s** \n Lowest Damage (1 run): **%s** \n Highest Damage (1 run): **%s** \n Lowest Combined Damage (2 runs): **%s** \n Highest Combined Damage (2 runs): **%s** \n Total Damage Dealt to Tyrfas (all teams): **%s**" % (gc1RunsCompleted, gc1TotalRuns, gc1Lowest, gc1Highest, gc1LowestCombined, gc1HighestCombined, gc1TotalDamageDone) 
    gc2String = "Total Runs Completed: **%s** / **%s** \n Lowest Damage (1 run): **%s** \n Highest Damage (1 run): **%s** \n Lowest Combined Damage (2 runs): **%s** \n Highest Combined Damage (2 runs): **%s** \n Total Damage Dealt to Lakreil (all teams): **%s**" % (gc2RunsCompleted, gc2TotalRuns, gc2Lowest, gc2Highest, gc2LowestCombined, gc2HighestCombined, gc2TotalDamageDone) 
    gc3String = "Total Runs Completed: **%s** / **%s** \n Lowest Damage (1 run): **%s** \n Highest Damage (1 run): **%s** \n Lowest Combined Damage (2 runs): **%s** \n Highest Combined Damage (2 runs): **%s** \n Total Damage Dealt to Velkazar (all teams): **%s**" % (gc3RunsCompleted, gc3TotalRuns, gc3Lowest, gc3Highest, gc3LowestCombined, gc3HighestCombined, gc3TotalDamageDone) 
    gcHistoricalString = "Historical Highest Damage Dealt to Tyrfas (single run): **%s** \n Historical Highest Damage Dealt to Lakreil (single run): **%s** \n Historical Highest Damage Dealt to Velkazar (single run): **%s**" % (gc1HistoricalHighest, gc2HistoricalHighest, gc3HistoricalHighest)
    gcDamageEmbed = discord.Embed(title="Damage (in billions) for the current Guild Conquest", color=0xFFFF00)
    gcDamageEmbed.add_field(name='Guild Conquest 1 (Tyrfas)', value=gc1String, inline=False)
    gcDamageEmbed.add_field(name='Guild Conquest 2 (Lakreil)', value=gc2String, inline=False)
    gcDamageEmbed.add_field(name='Guild Conquest 3 (Velkazar)', value=gc3String, inline=False)

    gcDamageEmbed.add_field(name='Historical Damage Records', value=gcHistoricalString, inline=False)

    await client.send_message(currentChannel, embed=gcDamageEmbed)

# async def retrieveDmg(user, message):

#     print ("Retrieving in Progress")
#     for i in range(0,50):
#         while True:
#             try:
#                 scope = ['https://spreadsheets.google.com/feeds',
#                          'https://www.googleapis.com/auth/drive']

#                 credentials = ServiceAccountCredentials.from_json_keyfile_name('PTBot.json', scope)

#                 gc = gspread.authorize(credentials)
#             except:
#                 print('Exception detected. retrying...')
#                 continue
#             break
#     sht1 = gc.open_by_key(googleAPIKey)
#     guildConquestSheet = sht1.get_worksheet(0)

#     runsDamageCells = [11, 14, 17, 28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 67, 70, 73, 76]
#     # runsDamageCells = [28, 31, 34, 37, 40, 43, 46, 49, 52, 55, 58, 61, 64, 67, 70, 73, 76]

#     for i, x in enumerate(runsDamageCells):

#         stringRange = 'I%s:J%s' % (x, x)
#         cell_list = guildConquestSheet.range(stringRange)
#         # print (cell_list)

#         # emptyFlag = 0
#         counter = 0
#         damageVal = []

#         for y in cell_list:
#             if y.value == '':
#                 damageVal.append('0')
#             elif y.value != '':
#                 damageVal.append(y.value)

#         # print (damageVal)
#         nameRange = 'E%s:E%s' % (str(x), str(x+2))
#         actualRange = 'F%s:F%s' % (str(x), str(x+2))

#         # print (nameRange)
#         # print (actualRange)
#         nameCells = guildConquestSheet.range(nameRange)
#         actualCells = guildConquestSheet.range(actualRange)

#         print (nameCells)
#         print (actualCells)

#         finalString = None

#         # print (combinedCells)

#         if '0' not in damageVal:
#             for j, z in enumerate(nameCells):
#                 if nameCells[j].value != '' and actualCells[j].value != '':
#                     # finalString = "%s (%s), " + finalString % (nameCells[j].value, actualCells[j].value) + finalString
#                     finalString = nameCells[j].value + '(' + actualCells[j].value + ')'
#                     print (finalString)
#                     # embed.add_field(name=formatString, value='test', inline=True)

#             embed = discord.Embed(title="Group %s" % (i+1), description=finalString, color=0xf20b0b)

#             embed.add_field(name='Run 1', value=damageVal[0], inline=True)
#             embed.add_field(name='Run 2', value=damageVal[1], inline=True)

#             await client.send_message(message.channel, embed=embed)

async def updateRoles():
    global PTMemberArray
    # global PTMemberArraywID
    global PTTrialArray
    global PTModArray
    # global scope
    # global credentials
    # global gc
    # global sht1
    # global guildConquestSheet
    # global worksheetSettings
    # global discordID

    await client.wait_until_ready()
    while not client.is_closed:

        # print ("Updating PT Roles...")
        server = client.get_server(id=serverID)
        tempPTMemberArray = []
        # tempPTMemberArraywID = []
        tempPTTrialArray = []
        tempPTModArray = []
        if server:
            for member in server.members:
                for roles in member.roles:
                    if str(roles) == 'PT Member':
                        tempPTMemberArray.append(member.id)
                        # tempPTMemberArraywID.append([member.name, member.id])
                    if str(roles) == 'PT Trial':
                        tempPTTrialArray.append(member.id)
                    if str(roles) == 'Mod':
                        tempPTModArray.append(member.id)

        PTMemberArray = tempPTMemberArray
        # PTMemberArraywID = tempPTMemberArraywID
        PTTrialArray = tempPTTrialArray
        PTModArray = tempPTModArray

        # rolesJSON = {"PTMember": PTMemberArray, "PTTrial":PTTrialArray, "Mod":PTModArray}
        # print (json.dumps(rolesJSON, indent=4))
        await asyncio.sleep(1800) # task runs every 1800 seconds

def create_assertion_session(conf_file, scopes, subject=None):
    with open(conf_file, 'r') as f:
        conf = json.load(f)

    token_url = conf['token_uri']
    issuer = conf['client_email']
    key = conf['private_key']
    key_id = conf.get('private_key_id')

    header = {'alg': 'RS256'}
    if key_id:
        header['kid'] = key_id

    # Google puts scope in payload
    claims = {'scope': ' '.join(scopes)}
    return AssertionSession(
        grant_type=AssertionSession.JWT_BEARER_GRANT_TYPE,
        token_url=token_url,
        issuer=issuer,
        audience=token_url,
        claims=claims,
        subject=subject,
        key=key,
        header=header,
    )

@client.event
async def on_ready():

    global authenticationFlag
    global authenticationFailCount

    print ("Authentication in progress...")

    # authWithAuthLib();
    while authenticationFlag == 0:
        await client.send_message(discord.Object(id=botCommandsID), 'Bot is online, awaiting authentication...')
        # print("1")
        if authWithAuthLib() == 1:
           await client.send_message(discord.Object(id=botCommandsID), 'Authentication successful.')
           # print("2")
           authenticationFlag = 1
        else:
           authenticationFailCount = authenticationFailCount + 1

           if authenticationFailCount != 3:
            await client.send_message(discord.Object(id=botCommandsID), 'Authentication unsuccessful, retrying in 5 seconds... (' + str(authenticationFailCount) + '/3 retries)')
            # print("3")
           else:
            await client.send_message(discord.Object(id=botCommandsID), 'Authentication unsuccessful, system exit initiated (' + str(authenticationFailCount) + '/3 retries)')
            # print("4")
            sys.exit(1)
           authenticationFlag = 0
           
           time.sleep(5)

    print ("------------------------------------")
    print ("Bot Name: " + client.user.name)
    print ("Bot ID: " + client.user.id)
    print ("Discord Version: " + discord.__version__)
    print ("------------------------------------")

    print ("Authentication successful.")
    # workbook = xlsxwriter.Workbook('arrays.xlsx')
    # worksheet = workbook.add_worksheet()

    # row = 0

    # for col, data in enumerate(PTMemberArraywID):
    #     print (col)
    #     print (data)
    #     worksheet.write_row(col, row, data)

    # workbook.close()

# client.loop.create_task(auth())

client.loop.create_task(updateRoles())
client.run(discordToken)