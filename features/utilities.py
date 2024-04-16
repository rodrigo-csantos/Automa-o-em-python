import shutil
import os
import PyPDF2

medicamentosRefrigerados = [
   "Actemra", "ACTEMRA",
   "Actemra SC", "ACTEMRA SC",
   "Adcetris", "ADCETRIS",
   "Alkeran", "ALKERAN",
   "Avastin", "AVASTIN",
   "Bavencio", "BAVENCIO",
   "Benlysta", "BENLYSTA",
   "Cetrotide", "CETROTIDE",
   "Cimzia", "CIMZIA",
   "Copaxone", "COPAXONE",
   "Cosentyx", "COSENTYX",
   "Ddavp Injetável", "DDAVP INJETÁVEL",
   "Doxopeg", "DOXOPEG",
   "Dupixent", "DUPIXENT",
   "Dysport", "DYSPORT",
   "Eligard", "ELIGARD",
   "Elonva", "ELONVA",
   "Emend", "EMEND",
   "Enbrel", "ENBREL",
   "Enbrel PFS", "ENBREL PFS",
   "Entyvio", "ENTYVIO",
   "Eprex", "EPREX",
   "Erbitux", "ERBITUX",
   "Euflexxa", "EUFLEXXA",
   "Eylia", "EYLIA",
   "Fabrazyme", "FABRAZYME",
   "Fasenra", "FASENRA",
   "Faslodex", "FASLODEX",
   "Faulblastina", "FAULBLASTINA",
   "Fauldacar", "FAULDACAR",
   "Fauldcita", "FAULDCITA",
   "Fauldleuco", "FAULDLEUCO",
   "Fauldmetro", "FAULDMETRO",
   "Fauldoxo", "FAULDOXO",
   "Fauldvincri", 'FAULDVINCRI',
   "Sulfato de Vincristina", "SULFATO DE VINCRISTINA",
   "Forteo", "FORTEO",
   "Gazyva Vials", "GAZYVA VIALS",
   "Genotropin (C5)", "GENOTROPIN (C5)",
   "Genotropin Goquick (C5)", "GENOTROPIN GOQUICK (C5)"
   "Glucagen Hypokit", "GLUCAGEN HYPOKIT",
   "Gonal F Pen", "GONAL F PEN",
   "Gonapeptyl Daily", "GONAPEPTYL DAILY",
   "Gonapeptyl Depot", "GONAPEPTYL DEPOT",
   "Granulokine", "GRANULOKINE",
   "Hemcibra Vials", "HEMCIBRA VIALS",
   "Herceptin", "HERCEPTIN",
   "Herceptin Sc", "HERCEPTIN SC",
   "Humalog", "HUMALOG",
   "Humira", "HUMIRA",
   "Humira Pen", "HUMIRA PEN",
   "Humira Vial", "HUMIRA VIAL"
   "Ilaris", "ILARIS",
   "Imfinzi", "IMFINZI",
   "Imunoglobulin", "IMUNOGLOBULIN",
   "Kadcyla", "KADCYLA",
   "Keytruda", "KEYTRUDA",
   "Kyprolis", "KYPROLIS",
   "Lantus", "LANTUS",
   "Lantus Solostar", "LANTUS SOLOSTAR",
   "Lectrum", "LECTRUM",
   "Lemtrada", "LEMTRADA",
   "Leukeran", "LEUKERAN",
   "Levemir Flexpen", "LEVEMIR FLEXPEN",
   "Levemir Penfil", "LEVEMIR PENFIL",
   "Lucentis", "LUCENTIS",
   "Lupron Kit", "LUPRON KIT",
   "Mabthera", "MABTHERA",
   "Mabthera Sc", "MABTHERA SC",
   "Mekinist", "MEKINIST",
   "Menopur Md", "MENOPUR MD",
   "Merional", "MERIONAL",
   "Mevatyl", "MEVATYL",
   "Navelbine", "NAVELBINE",
   "Neulastim", "NEULASTIM",
   "Norditropin Flexpro", "NORDITROPIN FLEXPRO",
   "Novomix", "NOVOMIX",
   "Novorapid", "NOVORAPID",
   "Nucala", "NUCALA",
   "Ocrevus", "OCREVUS",
   "Ofev", "OFEV"
   "Omnitrope", "OMNITROPE",
   "Orencia Sc", "ORENCIA SC",
   "Ovidrel", "OVIDREL",
   "Ozempic", "OZEMPIC",
   "Pergoveris Pen", "PERGOVERIS PEN",
   "Perjeta", "PERJETA",
   "Poemmy", "POEMMY",
   "Prolia", "PROLIA",
   "Puregon Pen", "PUREGON PEN",
   "Rebif", "REBIF",
   "Rekovelle", "REKOVELLE",
   "Remicade", "REMICADE",
   "Repatha", "REPATHA",
   "Saizen Liq", "SAIZEN LIQ",
   "Sandostatin", "SANDOSTATIN",
   "Saxenda", "SAXENDA",
   "Simponi", "SIMPONI",
   "Simulect", "SIMULECT",
   "Stelara", "STELARA",
   "Synagis", "SYNAGIS",
   "Taltz", "TALTZ",
   "Tecentriq", "TECENTRIQ",
   "Toujeo", "TOUJEO",
   "Tremfya", "TREMFYA",
   "Tresiba Flextouch", "TRESIBA FLEXTOUCH",
   "Tresiba Penfill", "TRESIBA PENFILL",
   "Trulicity", "TRULICITY",
   "Tysabri", "TYSABRI",
   "Victoza", "VICTOZA",
   "Vyndaqel", "VYNDAQEL",
   "Xgeva", "XGEVA",
   "Xolair", "XOLAIR",
   "Xultophy", "XULTOPHY",
   "Zedora", "ZEDORA",
   "Zoladex", "ZOLADEX",
   "Tecnocris", "TECNOCRIS",
   "N-Plate", "N-PLATE",
   "Riximyo", "RIXIMYO",
   "Vivaxxia", "VIVAXXIA",
   "Skyrizi", "SKYRIZI",
   "Tecentriq", "TECENTRIQ",
   "Libtayo", "LIBTAYO",
   "Pasurta", "PASURTA",
   "Emgality", "EMGALITY",
   "Afrezza", "AFREZZA",
   "Hyrimoz", "HYRIMOZ",
   "Takhzyro","TAKHZYRO",
   "Cubicin", "CUBICIN",
   "Zerbaxa","ZERBAXA",
   "Xilfya", "XILFYA",
   "Erelzi", "ERELZI",
   "Hemcibra Vials", "HEMCIBRA VIALS",
   "Evrysdi", "EVRYSDI",
   "Evenity", "EVENITY",
   "Vsiqq", "VSIQQ",
   "Enhertu", "ENHERTU",
   "Evusheld", "EVUSHELD",
   "Jemperli", "JEMPERLI",
   "Renahavis", "RENAHAVIS",
   "Sportvis", "SPORTVIS",
   "Criscy", "CRISCY",
   "Hadlima", "HADLIMA",
   "Genryzon", "GENRYZON",
   "Ecalta", "ECALTA",
   "Xenpozyme", "XENPOZYME",
   "Remsima", "REMSIMA",
   "Ropolivy", "ROPOLIVY",
   "Imjudo", "IMJUDO",
   "Kyntheum", "KYNTHEUM",
   "Genuxal", "GENUXAL",
   "Ajovy", "AJOVY",
   "Fiasp", "FIASP",
   "Elovie", "ELOVIE",
   "Saphnelo", "SAPHNELO",
   "Botox", "BOTOX",
   "Fauldfluor","FAULDFLUOR",
   "Alfaepoetina Humana Recombinante","ALFAEPOETINA HUMANA RECOMBINANTE",
   "Eritromax", "ERITROMAX",
   "Filgrastine","FILGRASTINE",
   "Alfaepoetina", "ALFAEPOETINA",
   "Permese", "PERMESE",
   "Fauldcarbo", "FAULDCARBO",
   "Fauldcispla","FAULDCISPLA",
   "Botulift","BOTULIFT",
   "Somatuline", "SOMATULINE"
]

# Caminho para a pasta onde os arquivos Excel estão localizados
pasta_origem = 'Caminho da pasta de origem'

# Caminho para a pasta de destino
pasta_destino = 'Caminho da pasta de destino'
pasta_destinoPDF = 'Caminho da pasta de destino'

def moveFile(sourceFolder,destinationFolder, typeFile, fileName=None):
    if typeFile == ".pdf":
       files = os.listdir(sourceFolder)
       for file in files:
          if file.endswith(typeFile) and file == fileName:
            caminhoOrigem = os.path.join(sourceFolder, file)
            caminhoDestino = os.path.join(destinationFolder, file)
            shutil.move(caminhoOrigem, caminhoDestino)
    else:
      files = os.listdir(sourceFolder)
      for file in files:
         if file.lower().endswith(typeFile):
               filePath = os.path.join(sourceFolder, file)
               shutil.move(filePath, destinationFolder)
               break  
     
def renameFile(destinationFolder, newName, counter, typeFile):
   files = os.listdir(destinationFolder)
   c = 0
   for file in files:
      if file.lower().endswith(typeFile):
         c += 1 
         if c == counter:
            filePath = os.path.join(destinationFolder, file)
            filePath2 = os.path.join(destinationFolder, newName)
            os.rename(filePath, filePath2)
            break 

def filterRefrigeratedPDF(nomeArquivo):
   
   pdfFile = open(nomeArquivo, 'rb') 
   dadosPDF = PyPDF2.PdfReader(pdfFile)
   
   textoPDF = ''
   for numPage in range(len(dadosPDF.pages)):
      paginaPDF = dadosPDF.pages[numPage]
      textoPDF += paginaPDF.extract_text()

   for medicamento in medicamentosRefrigerados:
      if medicamento in textoPDF:
         return True
   
   return False

def filterCompany (nomeArquivo):
   pdfFile = open(nomeArquivo, 'rb') 
   dadosPDF = PyPDF2.PdfReader(pdfFile)
   
   textoPDF = ''
   for numPage in range(len(dadosPDF.pages)):
      paginaPDF = dadosPDF.pages[numPage]
      textoPDF += paginaPDF.extract_text()

   if 'Empresa1' in textoPDF or 'Empresa1' in textoPDF:
      return 'Empresa1'
   elif 'Empresa2' in textoPDF or 'Empresa2' in textoPDF:
      return 'Empresa2'
   elif 'Empresa3' in textoPDF or 'Empresa3' in textoPDF:
      return 'Empresa3'
   elif 'Empresa4' in textoPDF or 'Empresa4' in textoPDF:
      return 'Empresa4'
   else:
      return 'Verificar empresa'
      
def deleteFiles (destinationFolder):
   files = [f for f in os.listdir(destinationFolder) if f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.pdf')]
   for file in files:
      try:
        filePath = os.path.join(destinationFolder, file)
        os.remove(filePath)
        print(f"Arquivo {file} excluído com sucesso.")
      except Exception as e:
        print(f"Erro ao excluir {file}: {e}")