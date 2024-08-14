import anvil.secrets
import anvil.server
import csv
import pandas as pd
import json
import anvil.media

from anvil.tables import app_tables
from marshmallow import Schema, fields, post_load
from pprint import pprint
from datetime import date
from langchain_core.messages import HumanMessage
from langchain_core.prompts import ChatPromptTemplate
from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_community.document_loaders import OutlookMessageLoader
from langchain_community.chat_models import ChatOpenAI
from langchain_community.llms import openai
from langchain.chains.summarize import load_summarize_chain
from langchain.docstore.document import Document
from langchain.text_splitter import RecursiveCharacterTextSplitter





clientDict = ""
aliasesList = ""
matterList = ""

def ClientMatter(subject):
    clientDict = ""
    aliasesList = ""
    matterList = ""

    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key="sk-proj-mjyh3qYIN6Gi7k8vLLoNT3BlbkFJl3kHTGkvx22x0Tvsvtfr")
    AliasesString = aliasesList.__str__()

    template_string = """
    You are an attorney billing expert. Your job is to infer the client alias from the folowing subject line of an email: {text}       
    In most cases, the subject of the email will contain the alias. In those cases you will compare the email subject with the LIST OF APPROVED ALIASES and return the alias FROM THE LIST OF APPROVED ALIASES that best matches the subject line.   
    
    For example, where the email subject is "topix -- quick questions", you would compare this to the list of approved aliases and infer that Page v. Topix Pharmaceuticals is the best fit for the alias because none of the other aliases contain the work topix.
    
    In some cases, the subject will not contain enough information to infer the alias. In those case, you will look at the body of the email for information that matches an alias FROM THE LIST OF APPROVED ALIASES. 
    Example 1: you may infer the  alias: Page v. Topix Pharmaceuticals where the body of the email refers to a person named Page. 
    Example 2: you may infer the alias: Aguilera v. Turner Systems, Inc. where the subject of the email refers to Turner.

    IMPORTANT: YOUR RESPONSE SHOULD NOT BE CONVERSATIONAL. YOUR REPSONSE ONLY CONTAIN THE ALIAS FROM THE LIST OF APPROVED ALIASES WITHOUT ANY ADDITIONAL WORDS. 

    INCORRECT response: Based on the information provided in the email, the inferred client alias is "Gonzalez v. DS Electric, Inc."
    CORRECT response: Gonzalez v. DS Electric, Inc.

    IMPORTANT: IF YOU CANNOT INFER AN ALIAS FROM THE CONTENT PROVIDED, YOUR MUST RESPOND WITH THE SINGLE WORD: None

    INCORRECT response: The inferred client alias from the email is None.
    CORRECT response: none

    THE LIST OF APPROVED ALIASES FOLLOWS =  
        Aguilera v. Turner Systems, Inc.
        Alvarez v. Command Security Services
        Barnwell v. Gilton Solid Waste Management
        Bhatia v. Mojio
        Buksh v. Sixt Rent A Car
        CAB v. UNITED SITE SERVICES, INC.
        Casbeer v. Friends of Downtown SLO
        Castellanos v. Urner’s, Inc.
        Clement v. Rescue Mission Alliance
        Cooper Aerial - PSG Contract Review
        Crowe v. Alternative Power Generation
        Deus v. Cuvaison, Inc.
        Diaz v. Smartlink
        Dufour v. Mann, Urrieta, & Nelson
        Ghasemi v .Valentia Analytical
        Gonzalez Davalos v. Nouveau Bakery LLC 
        Gonzalez v. DS Electric, Inc.
        Hale v. Belmont Village
        Hernandez v. Zarate Foods
        Hicks v. SSA Group, LLC
        Irmer v. KaiserAir, Inc.
        Jackson v. Mental Health Systems dba TURN
        Jermane v. Bethany Home Society of San Joaquin
        Jimenez v. Wade et al (Real Broker, LLC)
        Koeppen v. Cooper Aerial
        Kolkmann v. Alternative Power Generation
        Kumar DLSE De Novo Appeals
        McWillie v. Tenet Enterprises
        Melendez v. Peters Fruit Farms, Inc. 
        Miller v. Urata & Sons Concrete
        Nuno v. PCC Logistics
        Olson v. Allen Property Group, Inc.
        Oracle Anesthesia, Inc. v. Central Valley Anesthesia Partners
        Page v. Topix Pharmaceuticals
        Raj v. WFP Hospitality II LLC
        Rentner v. Trimble, Fletchers
        Rocio Tafoya v. Peters Fruit Farms, Inc.
        Sandoval v. The Nephrology Group, Inc.
        SBM v. Baker
        Spooner v. Tri-Ced Economic Development Corporation
        Staedler v. SSA Group, LLC
        Steve v. Tulare Firestone, Inc.
        Sweet Adeline v. Tasty Wings
        Thomas v. US Tech, Quest Media
        Turpin v. Sinclair Broadcast Group, Inc.
        Vaughn v. United Freight Lines
        Webb v. Doug Tauzer Construction
        White v. Andy’s Produce Market
        Wood v. Smartlink
    """


    prompt = PromptTemplate.from_template(template_string)
    
    chain = load_summarize_chain(llm, chain_type="stuff", prompt=prompt)

    
    doc =  Document(page_content=subject)
    output_clientmatter = chain.run([doc])
    output = output_clientmatter.strip()         
    return output

def Narrative(msg_recipient, msg_from, msg_body, msg_subject):
    llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo-16k", api_key="sk-proj-mjyh3qYIN6Gi7k8vLLoNT3BlbkFJl3kHTGkvx22x0Tvsvtfr")
    prompt_template = """
    You are a secretary working for attorney Daniel Cravens. Your job is to create a billing entry that succinctly summarizes the work that Daniel Cravens performed based on the email provided. You must begin your billing entry with a verb. 
    
    EXAMPLE 1: Where Daniel Cravens is emailing with a person outside of the ohaganmeyer.com domain, begin the billing entry with "Email communication with [insert name of person to whom Daniel was communicating] concerning [description of work]. 
    
    Example 2: Where Daniel Cravens is email the opposing attorney on the case, the summary would begin "Meet and confer correspondence with opposing counsel [insert name of opposing counsel] regarding [insert subject matter of discussion]"

    EXAMPLE 3: Where Daniel Cravens is providing instructions to a person within the ohaganmeyer.com domain, the work performed should be written work product that will ultimately be produced but should not mention the name of the people. For example, where the Daniel instructions Caleb to prepare a shell for a motion to compel, the entry would be: "Update and revise motion to compel"
    
    Email to summarize: "{text}"
    
    """
    prompt = PromptTemplate.from_template(prompt_template)
    chain = load_summarize_chain(llm, chain_type="stuff", prompt=prompt)
    stuffedshit = msg_subject + msg_from + msg_recipient + msg_body
    docs =  Document(page_content=stuffedshit)
    output_summary = chain.run([docs])
    print("Narrative: ", output_summary)
    return output_summary
