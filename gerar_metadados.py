# gerar_metadados.py

conteudo_versao = """VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'Philipe Sampaio (oldboy465)'),
         StringStruct(u'FileDescription', u'Extrator Avancado de GRS para UEMA'),
         StringStruct(u'FileVersion', u'1.0.0'),
         StringStruct(u'InternalName', u'ExtratorGRS'),
         StringStruct(u'LegalCopyright', u'Copyright (c) 2026 Philipe Sampaio. All rights reserved.'),
         StringStruct(u'OriginalFilename', u'extrator_uema.exe'),
         StringStruct(u'ProductName', u'Extrator de GRS - UEMA'),
         StringStruct(u'ProductVersion', u'1.0.0')])
      ]), 
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)"""

with open("version.txt", "w", encoding="utf-8") as f:
    f.write(conteudo_versao)

print("Arquivo 'version.txt' criado com sucesso! Agora você pode usar no PyInstaller.")