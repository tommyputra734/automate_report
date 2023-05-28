import discord
from discord import SyncWebhook


def send_to_discord(webhook_url, output_data):
    """Sending File alerts to discord

    Parameters
    ----------
    webhook_url : string
        _description_
    output_data : string
        _description_
    """
    webhook = SyncWebhook.from_url(webhook_url)

    with open(file=output_data, mode='rb') as file:
        excel_file = discord.File(file)

    webhook.send('This is an automated report', 
                username='Sales Bot', 
                file=excel_file)