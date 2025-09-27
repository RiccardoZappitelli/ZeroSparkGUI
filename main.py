import json
import pandas as pd
import customtkinter as ctk
from CTkScrollableDropdown import CTkScrollableDropdown

datasheet_filename = "ids.xlsx"
template_filename = "template.json"

dfs = pd.ExcelFile(datasheet_filename)
sheets = {sheet_name:dfs.parse(sheet_name) for sheet_name in dfs.sheet_names}

characters_sheet = sheets["Characters"]
characters_dict = {character[0]:character[1] for character in characters_sheet.values}
all_characters = list(characters_dict.keys())

label_placeholder = "CharacterLabel"
create_character_placeholder = f"CreateCharacter({label_placeholder})"
customize_character_placeholder = f"CustomizeCharacter({label_placeholder})"

#Characters, Stages, Always Aura, BGM

always_on_aura_keys = {
    "AlwaysAuraKey": "特殊常時オーラ", "BattleAlwaysAuraKey": "特殊常時オーラ",
}
#"SetAudioFiles"
set_audio = {
    "TargetCharacterID": "",
    "SFX": "/Game/SS/Sounds/Battle_SE/BTLSE_IDPLACEHOLDER",
    "English": "/Game/SS/Sounds/Battle_VOICE/US/BTLCV_IDPLACEHOLDER_US",
    "Japanese": "/Game/SS/Sounds/Battle_VOICE/US/BTLCV_IDPLACEHOLDER_JP"
}

with open(template_filename, "r") as fi:
    template = json.load(fi)

class GUI:
    def __init__(self) -> None:
        self.root = ctk.CTk()
        self.root.title("ZeroSpark Json Maker")
        self.root.geometry("900x500")
        self.root.resizable(True, False)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)


        #CHARACTER CREATION
        self.characterCreationFrame = ctk.CTkFrame(self.root)
        self.characterCreationFrame.grid(row=0, column=0, padx=5, pady=5)

        self.characterCreationFrameLabel = ctk.CTkLabel(self.characterCreationFrame, text="Character Creation")
        self.characterCreationFrameLabel.grid(row=0, column=0, sticky="nsew", columnspan=4)

        self.characterLabelLabel = ctk.CTkLabel(self.characterCreationFrame, text="Label")
        self.characterLabelLabel.grid(row=1, column=0, pady=12, padx=12)

        self.characterLabelEntry = ctk.CTkEntry(self.characterCreationFrame)
        self.characterLabelEntry.grid(row=1, column=1, pady=12, padx=12)

        self.characterNameLabel= ctk.CTkLabel(self.characterCreationFrame, text="Name")
        self.characterNameLabel.grid(row=1, column=2, pady=12, padx=12)

        self.characterNameEntry = ctk.CTkEntry(self.characterCreationFrame) 
        self.characterNameEntry.grid(row=1, column=3, pady=12, padx=12)

        self.formNameLabel= ctk.CTkLabel(self.characterCreationFrame, text="FormName")
        self.formNameLabel.grid(row=2, column=0, pady=12, padx=12)

        self.formNameEntry = ctk.CTkEntry(self.characterCreationFrame)
        self.formNameEntry.grid(row=2, column=1, pady=12, padx=12)

        self.baseCharacterIDLabel= ctk.CTkLabel(self.characterCreationFrame, text="BaseCharacterID")
        self.baseCharacterIDLabel.grid(row=2, column=2, pady=12, padx=12)

        self.baseCharacterIDEntry = ctk.CTkEntry(self.characterCreationFrame) 
        self.baseCharacterIDEntry.grid(row=2, column=3, pady=12, padx=12)
        CTkScrollableDropdown(attach=self.baseCharacterIDEntry, values=all_characters, width=300, autocomplete=True)

        self.newCharacterIDLabel= ctk.CTkLabel(self.characterCreationFrame, text="NewCharacterID")
        self.newCharacterIDLabel.grid(row=3, column=0, pady=12, padx=12)

        self.newCharacterIDEntry = ctk.CTkEntry(self.characterCreationFrame) 
        self.newCharacterIDEntry.grid(row=3, column=1, pady=12, padx=12)

        self.rosterPositionLabel = ctk.CTkLabel(self.characterCreationFrame, text="Roster Position")
        self.rosterPositionLabel.grid(row=3, column=2, pady=12, padx=12)

        self.rosterPositionEntry = ctk.CTkEntry(self.characterCreationFrame)
        self.rosterPositionEntry.grid(row=3, column=3, pady=12, padx=12)
        CTkScrollableDropdown(attach=self.rosterPositionEntry, values=all_characters, width=300)

        #CHARACTER CUSTOMIZATION
        self.characterCustomizationFrame = ctk.CTkFrame(self.root)
        self.characterCustomizationFrame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        self.characterCustomizationFrameLabel = ctk.CTkLabel(self.characterCustomizationFrame, text="Character Customization")
        self.characterCustomizationFrameLabel.grid(row=0, column=0, sticky="nsew", columnspan=4)

        self.aurafxLabel = ctk.CTkLabel(self.characterCustomizationFrame, text="AuraFX")
        self.aurafxLabel.grid(row=1, column=0, padx=12, pady=12)

        self.aurafxEntry = ctk.CTkEntry(self.characterCustomizationFrame) 
        self.aurafxEntry.grid(row=1, column=1, padx=12, pady=12)
        CTkScrollableDropdown(attach=self.aurafxEntry, values=all_characters, width=150)

        self.auraAlwaysOnLabel = ctk.CTkLabel(self.characterCustomizationFrame, text="AuraAlwaysOn")
        self.auraAlwaysOnLabel.grid(row=1, column=2, padx=12, pady=12)

        self.auraAlwaysOnEntry = ctk.CTkSwitch(self.characterCustomizationFrame, text="", command=lambda:...)
        self.auraAlwaysOnEntry.grid(row=1, column=3, padx=12, pady=12)

        self.skill1Label = ctk.CTkLabel(self.characterCustomizationFrame, text="Skill1")
        self.skill1Label.grid(row=2, column=0, padx=12, pady=12)

        self.skill1Entry = ctk.CTkEntry(self.characterCustomizationFrame) 
        self.skill1Entry.grid(row=2, column=1, padx=12, pady=12)
        CTkScrollableDropdown(attach=self.skill1Entry, values=all_characters, width=150)

        self.skill2Label = ctk.CTkLabel(self.characterCustomizationFrame, text="Skill2")
        self.skill2Label.grid(row=2, column=2, padx=12, pady=12)

        self.skill2Entry = ctk.CTkEntry(self.characterCustomizationFrame) 
        self.skill2Entry.grid(row=2, column=3, padx=12, pady=12)
        CTkScrollableDropdown(attach=self.skill2Entry, values=all_characters, width=150)

        self.blast1Label = ctk.CTkLabel(self.characterCustomizationFrame, text="Blast1")
        self.blast1Label.grid(row=3, column=0, padx=12, pady=12)

        self.blast1Entry = ctk.CTkEntry(self.characterCustomizationFrame) 
        self.blast1Entry.grid(row=3, column=1, padx=12, pady=12)
        CTkScrollableDropdown(attach=self.blast1Entry, values=all_characters, width=150)

        self.blast2Label = ctk.CTkLabel(self.characterCustomizationFrame, text="Blast2")
        self.blast2Label.grid(row=3, column=2, padx=12, pady=12)

        self.blast2Entry = ctk.CTkEntry(self.characterCustomizationFrame) 
        self.blast2Entry.grid(row=3, column=3, padx=12, pady=12)
        CTkScrollableDropdown(attach=self.blast2Entry, values=all_characters, width=150)

        self.ultimateLabel = ctk.CTkLabel(self.characterCustomizationFrame, text="Ultimate")
        self.ultimateLabel.grid(row=4, column=2, padx=12, pady=12)

        self.ultimateEntry = ctk.CTkEntry(self.characterCustomizationFrame) 
        self.ultimateEntry.grid(row=4, column=3, padx=12, pady=12)
        CTkScrollableDropdown(attach=self.ultimateEntry, values=all_characters, width=150)

        self.create_button = ctk.CTkButton(self.characterCustomizationFrame, text="Create", command=self.create)
        self.create_button.grid(row=5, column=0, padx=12, pady=12)

        self.characterCreationFrame.grid_columnconfigure((0,1,2,3), weight=1)
        self.characterCustomizationFrame.grid_columnconfigure((0,1,2,3), weight=1)

        self.root.mainloop()

    def getData(self) -> None:
        label = self.characterLabelEntry.get()
        name = self.characterNameEntry.get()
        formname = self.formNameEntry.get()

        basecharacterID = self.baseCharacterIDEntry.get()
        basecharacterID = characters_dict[basecharacterID]

        newcharacterID = self.newCharacterIDEntry.get()

        rosterposition = self.rosterPositionEntry.get()
        rosterposition = characters_dict[rosterposition]

        aurafx = self.aurafxEntry.get()
        aurafx = characters_dict[aurafx]

        always_on_aura = self.auraAlwaysOnEntry.get()

        skill1 = self.skill1Entry.get()
        skill1 = characters_dict[skill1]

        skill2 = self.skill2Entry.get()
        skill2 = characters_dict[skill2]

        blast1 = self.blast1Entry.get()
        blast1 = characters_dict[blast1]

        blast2 = self.blast2Entry.get()
        blast2 = characters_dict[blast2]

        ultimate = self.ultimateEntry.get()
        ultimate = characters_dict[ultimate]

        print(template)
        template[label] = template[label_placeholder]
        del template[label_placeholder]

        new_create_label = f"CreateCharacter({label})"
        template[label][new_create_label] = template[label].pop(create_character_placeholder)
        creation_layer = template[label][new_create_label]
        creation_layer["NewCharacterID"] = newcharacterID
        creation_layer["BaseCharacterID"] = basecharacterID
        creation_layer["Name"] = name
        creation_layer["FormName"] = formname
        creation_layer["SamePerson"] = basecharacterID
        creation_layer["CopyRosterPosition"] = rosterposition

        new_customize_label = f"CustomizeCharacter({label})"
        template[label][new_customize_label] = template[label].pop(customize_character_placeholder)
        customization_layer = template[label][new_customize_label]
        customization_layer["TargetCharacterID"] = newcharacterID
        customization_layer["AuraFX"] = aurafx
        customization_layer["Skill1"] = skill1
        customization_layer["Skill2"] = skill2
        customization_layer["Blast1"] = blast1
        customization_layer["Blast2"] = blast2
        customization_layer["Ultimate"] = ultimate
        if always_on_aura:
            customization_layer.update(always_on_aura_keys)

        set_audio["TargetCharacterID"] = newcharacterID
        set_audio["SFX"] = set_audio["SFX"].replace("IDPLACEHOLDER", ultimate)
        set_audio["English"] = set_audio["English"].replace("IDPLACEHOLDER", ultimate)
        set_audio["Japanese"] = set_audio["Japanese"].replace("IDPLACEHOLDER", ultimate)
        template[label].update({
            "SetAudioFiles":set_audio
        })
        return template, f"{label.capitalize()}.json"

    def create(self) -> None:
        data, filename = self.getData()
        with open(filename, "w", encoding="utf-8") as fo:
            json.dump(data, fo)

if __name__ == "__main__":
    gui=GUI()