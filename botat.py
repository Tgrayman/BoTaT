import os
import random

import discord
from dotenv import load_dotenv
from discord.ext import commands
import asyncio
import xlrd
import logging

# Initialisations et autres data
load_dotenv()
TOKEN = os.getenv('DISCORD_TOKEN')
GUILD = os.getenv('DISCORD_GUILD_NAME')
GUILD_ID = os.getenv('DISCORD_GUILD_ID')

description = """BoTaT - Attribution des numéros d'anonymat et changement de grade"""
botat = commands.Bot(command_prefix='$', description=description)

tat_ID = int(GUILD_ID)
server = botat.get_guild(tat_ID)
temp = range(225)
code_left = list(temp)
users_done = []

@botat.event
async def on_ready():
    game = discord.Game("Prêt à vérifier et distribuer")
    await botat.change_presence(status=discord.Status.online, activity=game)
    guild = discord.utils.get(botat.guilds, name=GUILD)
    print(f'{botat.user} is connected to the following guild:\n'
          f'{guild.name}(id: {guild.id})')


@botat.command()
async def ping(ctx):
    "Ping le bot"
    await ctx.send(f"*Pong, vitesse de {round(botat.latency * 1000)} ms*")


@botat.command(name='hello', help='un simple hello world')
async def hello(ctx):
    await ctx.send("Hello world")


@botat.command(name='cake', help= 'Portal')
async def cake(ctx):
    await ctx.send("The cake is a lie !")


@botat.command(name='code', help="distribue le code de vérification pour le vote")
async def validate(ctx, member: discord.Member):
    if member in users_done:
        await ctx.send("Erreur : cette personne a déjà reçu son code !")
        print("Error : This person has already a key")
    else:
        loc1 = "keys.xlsx"
        wb_rd = xlrd.open_workbook(loc1)
        sheet_read = wb_rd.sheet_by_index(0)
        try:
            i = random.randint(0, len(code_left))
            x = code_left[i]
            code = sheet_read.cell_value(x, 0)
            code_left.remove(x)
            users_done.append(member)
        except:
            await ctx.send("Erreur : plus de codes disponible !")
            print("No key remaining")
        else:
            await member.send(f'Votre Code pour voter est {code}\n'
                          f'Ce code est nécessaire pour que votre vote soit comptabilisé. Ne le divulguez pas où votre voix ne poourra compter')
            await ctx.send("Code envoyé")
            print('Code send')


@botat.command(name='codes_restants', help="commande servant à afficher le nombre de codes restants")
async def codes_restants(ctx):
    await ctx.send(f'il reste {len(code_left)} codes non distribués')
    print(f"{len(code_left)} codes left")
    print(code_left)


@botat.command(name='role_test',
               help="commande servant simplement à séparer le changement de rôle et la distribution du code pour plus "
                    "de facilité au développement")
async def role_test(ctx, member: discord.Member):
    role = discord.utils.get(server.roles, name='Membre actif')  # NOT WORKING
    if member.roles is not role:
        await member.add_roles(role)
        await ctx.send("rôle changé")
    else:
        await ctx.send("rôle non modifié")



botat.run(TOKEN)
