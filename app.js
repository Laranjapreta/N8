// Matriz de Risco — OFFLINE v8
const AGENT_COLORS = {"Físico":"#16a34a","Químico":"#dc2626","Ergonômico":"#f59e0b","Acidente":"#2563eb","Biológico":"#8B4513"};
const TPL = {
  contentTypes: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  workbook: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="29029"/><workbookPr defaultThemeVersion="202300"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="C:\\Users\\MT\\Downloads\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="8_{1B6F1BD2-76CA-4566-8D60-8DCCCC98EA0C}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/><bookViews><workbookView xWindow="-109" yWindow="-109" windowWidth="26301" windowHeight="14305" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews><sheets><sheet name="Riscos" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0"/></workbook>`,
  workbookRels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  styles: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><fonts count="5" x14ac:knownFonts="1"><font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font><font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/></font><font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDC2626"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE1EFED"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFB5D5D0"/></left><right style="thin"><color rgb="FFB5D5D0"/></right><top style="thin"><color rgb="FFB5D5D0"/></top><bottom style="thin"><color rgb="FFB5D5D0"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="10"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="4" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><colors><mruColors><color rgb="FFE1EFED"/><color rgb="FFF2F8F7"/><color rgb="FFB5D5D0"/></mruColors></colors><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`,
  theme1: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tema do Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6><a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont><a:latin typeface="Aptos Narrow" panose="02110004020202020204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults><a:lnDef><a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style></a:lnDef></a:objectDefaults><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{2E142A2C-CD16-42D6-873A-C26D2A0506FA}" vid="{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}"/></a:ext></a:extLst></a:theme>`,
  sharedStrings: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="42" uniqueCount="39"><si><t>Agente</t></si><si><t>Fator de Risco</t></si><si><t>Possível dano</t></si><si><t>Fund. legal / Limite de exposição</t></si><si><t>Fonte geradora</t></si><si><t>EPC</t></si><si><t>EPI</t></si><si><t>Atenuação</t></si><si><t>Intens./Concent.</t></si><si><t>Técnica</t></si><si><t>Tipo de exposição</t></si><si><t>Probabilidade</t></si><si><t>Severidade</t></si><si><t>Nível de risco</t></si><si><t>Químico</t></si><si><t>Produtos de limpeza (desengordurantes/hipoclorito/detergentes)</t></si><si><t>Dermatites; irritação respiratória/ocular</t></si><si><t>NR-26 (GHS) / FISPQ; NR-15 (Agentes químicos)</t></si><si><t>Detergentes</t></si><si><t>Ventilação geral</t></si><si><t>Luvas nitrílicas</t></si><si><t>NA</t></si><si><t>Avaliação Qualitativa</t></si><si><t>Habitual</t></si><si><t>Risco Baixo</t></si><si><t>Físico</t></si><si><t>Ruído</t></si><si><t>PAIR; zumbidos</t></si><si><t>NR-15 (Ruído)</t></si><si><t>Máquinas e Equipamentos</t></si><si><t>Protetor auricular</t></si><si><t>15 dB</t></si><si><t>86,5 db</t></si><si><t>NHO-01</t></si><si><t>Risco Médio</t></si><si><t>Pouco provável</t></si><si><t>Improvável</t></si><si><t>Leve</t></si><si><t>Sério</t></si></sst>`,
  sheetViews: `<sheetViews><sheetView tabSelected="1" zoomScale="70" zoomScaleNormal="70" workbookViewId="0"><selection activeCell="T7" sqref="T7"/></sheetView></sheetViews>`,
  cols: `<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="24.375" style="1" customWidth="1"/><col min="3" max="3" width="21.375" customWidth="1"/><col min="4" max="4" width="15.75" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/><col min="6" max="6" width="11.625" customWidth="1"/><col min="7" max="7" width="12.375" customWidth="1"/><col min="8" max="8" width="12.625" customWidth="1"/><col min="9" max="9" width="10.25" customWidth="1"/><col min="10" max="10" width="11.625" customWidth="1"/><col min="11" max="11" width="15" customWidth="1"/><col min="12" max="12" width="16.625" customWidth="1"/><col min="13" max="13" width="14.625" customWidth="1"/><col min="14" max="14" width="17.625" customWidth="1"/></cols>`,
  merges: ``,
};
const CATALOG_BASE={agentes:{"Físico":[{factor:"Calor",dano:"Estresse térmico",legal:"NR-15 (Calor)",fontes:["Fornos","Fogões","Chapas","Fritadeiras"]},{factor:"Ruído",dano:"PAIR; zumbidos",legal:"NR-15 (Ruído)",fontes:["Processos ruidosos","Exaustores","Equipamentos"]},{factor:"Vibração",dano:"Lesões osteomusculares; desconforto",legal:"NR-15 (Vibração)",fontes:["Ferramentas vibratórias"]},{factor:"Umidade",dano:"Dermatites; resfriados",legal:"NR-15 (Umidade)",fontes:["Lavagem constante","Ambiente úmido"]},{factor:"Frio",dano:"Hipotermia; desconforto térmico",legal:"NR-15 (Frio)",fontes:["Câmaras frias"]},{factor:"Iluminação inadequada",dano:"Fadiga visual; erros",legal:"NR-17 (Condições de conforto)",fontes:["Luminárias insuficientes","Ofuscamento"]}],"Químico":[{factor:"Produtos de limpeza (desengordurantes/hipoclorito/detergentes)",dano:"Dermatites; irritação respiratória/ocular",legal:"NR-26 (GHS) / FISPQ; NR-15 (Agentes químicos)",fontes:["Desengordurantes","Hipoclorito","Detergentes"]},{factor:"Solventes orgânicos (álcool/thinner)",dano:"Cefaleia; tontura; dermatites",legal:"NR-26 (GHS); NR-15 (Solventes)",fontes:["Desinfecção","Remoção de tintas"]},{factor:"Poeira orgânica (farinha)",dano:"Rinite; asma ocupacional",legal:"NR-15 (Poeiras)",fontes:["Panificação","Manuseio de farinha"]},{factor:"GLP (gás de cozinha)",dano:"Asfixia; incêndio/explosão",legal:"NR-20 (Inflamáveis)",fontes:["Cilindros e queimadores"]},{factor:"Cloro/Desinfetantes fortes",dano:"Queimaduras químicas; irritação respiratória",legal:"NR-26 (GHS) / FISPQ",fontes:["Desinfecção de pisos/superfícies"]}],"Ergonômico":[{factor:"Postura em pé prolongada",dano:"Lombalgias; fadiga postural",legal:"NR-17 (Ergonomia)",fontes:["Bancadas sem ajuste","Ritmo intenso"]},{factor:"Movimentos repetitivos",dano:"LER/DORT",legal:"NR-17 (Ergonomia)",fontes:["Corte repetitivo","Embalagem"]},{factor:"Levantamento/transporte de cargas",dano:"Entorses; hérnias",legal:"NR-17 (Ergonomia)",fontes:["Sacos de mantimentos"]},{factor:"Pausas insuficientes/ritmo acelerado",dano:"Estresse; fadiga",legal:"NR-17 (Organização do trabalho)",fontes:["Picos de demanda"]}],"Acidente":[{factor:"Cortes",dano:"Lacerações; cortes profundos",legal:"NR-12 / NR-01",fontes:["Facas","Fatiadores"]},{factor:"Queimaduras",dano:"Queimaduras de 1º a 3º",legal:"NR-23 / NR-01",fontes:["Óleo quente","Fritadeiras"]},{factor:"Queda do mesmo nível",dano:"Contusões; fraturas leves",legal:"NR-01",fontes:["Piso molhado/oleoso"]},{factor:"Incêndio/Explosão",dano:"Queimaduras graves; morte",legal:"NR-20 / NR-23",fontes:["GLP","Óleo quente"]},{factor:"Choque elétrico",dano:"Queimaduras; parada cardíaca",legal:"NR-10",fontes:["Tomadas danificadas","Extensões improvisadas"]}],"Biológico":[{factor:"Contato com alimentos crus (carnes/ovos)",dano:"Infecções alimentares; contaminação cruzada",legal:"NR-32 (Agentes biológicos) — quando aplicável; boas práticas/ANVISA",fontes:["Manuseio de carnes cruas","Ovos crus","Superfícies contaminadas"]},{factor:"Fungos/Mofos em áreas úmidas",dano:"Rinite; alergias; infecções respiratórias",legal:"NR-32 (Agentes biológicos) — quando aplicável",fontes:["Áreas com infiltração","Estocagem úmida"]},{factor:"Resíduos orgânicos",dano:"Agentes patogênicos diversos",legal:"NR-32 — quando aplicável; manejo de resíduos",fontes:["Lixo orgânico","Restos de alimentos"]}]},probOpts:[{value:1,label:"Possível"},{value:2,label:"Improvável"},{value:3,label:"Pouco provável"},{value:4,label:"Provável"}],sevOpts:[{value:1,label:"Leve"},{value:2,label:"Moderado"},{value:3,label:"Sério"},{value:4,label:"Incapacitante"}],tiposExposicao:["Habitual","Permanente","Eventual"],tecnicas:["Avaliação Qualitativa"],epc:["NA","Exaustão localizada","Ventilação geral","Barreiras/guardas","Tapete antiderrapante","Sinalização/Delimitação","Chuveiro lava-olhos","Extintores","Detecção de GLP","Isolamento térmico","Cortina de ar quente","Organização 5S"],epi:["NA","Luvas nitrílicas","Luvas de malha de aço (anticorte)","Luvas térmicas","Avental térmico","Avental impermeável","Óculos de proteção","Face shield","Respirador PFF2/N95","Protetor auricular","Botina antiderrapante","Creme de proteção"]};
const LS_KEYS={header:"sst_header",linhas:"sst_riscos_linhas",catalogUser:"sst_catalog_user_v2"};
function load(k,f){try{const v=localStorage.getItem(k);return v?JSON.parse(v):f}catch(e){return f}}
function save(k,v){localStorage.setItem(k,JSON.stringify(v))}
function getUser(){return load(LS_KEYS.catalogUser,{agentes:{}, tecnicas:[], tiposExposicao:[], epc:[], epi:[]})}
function setUser(u){save(LS_KEYS.catalogUser,u)}
function getCatalog(){const user=getUser();const merged=JSON.parse(JSON.stringify(CATALOG_BASE));Object.keys(user.agentes||{}).forEach(agent=>{if(!merged.agentes[agent]) merged.agentes[agent]=[];(user.agentes[agent]||[]).forEach(f=> merged.agentes[agent].push(f))}); merged.tecnicas=[...merged.tecnicas,...(user.tecnicas||[])]; merged.tiposExposicao=[...merged.tiposExposicao,...(user.tiposExposicao||[])]; merged.epc=[...merged.epc,...(user.epc||[])]; merged.epi=[...merged.epi,...(user.epi||[])]; return merged;}
const RISK_MATRIX={1:{1:{nivel:"Risco Irrelevante",cls:"sky"},2:{nivel:"Risco Baixo",cls:"green"},3:{nivel:"Risco Baixo",cls:"green"},4:{nivel:"Risco Médio",cls:"yellow"}},2:{1:{nivel:"Risco Baixo",cls:"green"},2:{nivel:"Risco Baixo",cls:"green"},3:{nivel:"Risco Médio",cls:"yellow"},4:{nivel:"Risco Alto",cls:"red"}},3:{1:{nivel:"Risco Baixo",cls:"green"},2:{nivel:"Risco Médio",cls:"yellow"},3:{nivel:"Risco Alto",cls:"red"},4:{nivel:"Risco Alto",cls:"red"}},4:{1:{nivel:"Risco Médio",cls:"yellow"},2:{nivel:"Risco Alto",cls:"red"},3:{nivel:"Risco Alto",cls:"red"},4:{nivel:"Risco Crítico",cls:"black"}}};
function matrixResult(p,s){return (RISK_MATRIX[p]?.[s])||RISK_MATRIX[1][1]}
function esc(s){return String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]))}

// INDEX
const STATE={header:load(LS_KEYS.header,{ghe:"01",setor:"COZINHA",fase:"Antecipação",cargos:"AUXILIAR DE COZINHA; COZINHEIRA",descricao:"Auxiliam nos serviços de alimentação..."}),linhas:load(LS_KEYS.linhas,[])};
function mountIndex(){
  document.body.innerHTML=`
  <header class="topbar">
    <h1>Matriz de Risco</h1>
    <div class="actions">
      <a class="btn primary" href="./cadastro-registro.html">+ Adicionar registro</a>
      <a class="btn" href="./cadastro-fator.html">+ Cadastrar novo fator de risco</a>
      <button class="btn" id="btn-xlsx">Exportar XLSX</button>
      <button class="btn danger" id="btn-clear">Limpar tudo</button>
    </div>
  </header>
  <main class="container">
    <section class="card">
      <h2>Identificação do GHE</h2>
      <div class="grid-4">
        <label>GHE <input id="ghe" value="${esc(STATE.header.ghe)}"></label>
        <label>Setor <input id="setor" value="${esc(STATE.header.setor)}"></label>
        <label>Fase <input id="fase" value="${esc(STATE.header.fase)}"></label>
        <label>Cargos (separe por ;) <input id="cargos" value="${esc(STATE.header.cargos)}"></label>
      </div>
      <label>Descrição das atividades <textarea id="descricao" rows="3">${esc(STATE.header.descricao)}</textarea></label>
    </section>
    <section class="card">
      <h2>Tabela de riscos do GHE</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Agente</th><th>Fator de Risco</th><th>Possível dano</th><th>Fund. legal / Limite de exposição</th>
              <th>Fonte geradora</th><th>EPC</th><th>EPI</th><th>Atenuação</th><th>Intens./Concent.</th>
              <th>Técnica</th><th>Tipo de exposição</th><th class="nowrap">Probabilidade</th><th class="nowrap">Severidade</th><th>Nível de risco</th>
            </tr>
          </thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
      <div class="actions">
        <a class="btn" href="./cadastro-registro.html">Adicionar registro</a>
        <a class="btn" href="./cadastro-fator.html">Cadastrar novo fator de risco</a>
      </div>
    </section>
  </main>`;
  ["ghe","setor","fase","cargos","descricao"].forEach(id=>{
    const el=document.getElementById(id);
    el.addEventListener("input",()=>{STATE.header[id]=el.value; save(LS_KEYS.header, STATE.header);});
  });
  document.getElementById("btn-clear").onclick=()=>{ if(confirm("Tem certeza?")){ STATE.linhas=[]; save(LS_KEYS.linhas,STATE.linhas); renderTable(); } };
  document.getElementById("btn-xlsx").onclick=exportXLSX;
  renderTable();
}
function renderTable(){
  const tb=document.getElementById("tbody"); tb.innerHTML="";
  if(!STATE.linhas.length){
    const tr=document.createElement("tr"); tr.className="empty";
    const td=document.createElement("td"); td.colSpan=14; td.textContent="Nenhum risco cadastrado ainda."; tr.appendChild(td); tb.appendChild(tr); return;
  }
  STATE.linhas.forEach((l,i)=>{
    const tr=document.createElement("tr"); tr.dataset.idx=i;
    const td=(html,cls='')=>{const t=document.createElement("td"); t.innerHTML=html; if(cls) t.className=cls; tr.appendChild(t);};
    td(`<span class="agent-pill"><span class="dot" style="background:${AGENT_COLORS[l.agente]||'#94a3b8'}"></span>${esc(l.agente)}</span>`);
    td(esc(l.fator)); td(esc(l.dano)); td(esc(l.legal)); td(esc(l.fonteGeradora||"-")); td(esc(l.epc||"-")); td(esc(l.epi||"-"));
    td(esc(l.atenuacao||"-")); td(esc(l.intensidade||"-")); td(esc(l.tecnica||"-")); td(esc(l.tipoExpo||"-"));
    td(`${l.prob} — ${esc(l.probLabel||"")}`,'center nowrap'); td(`${l.sev} — ${esc(l.sevLabel||"")}`,'center nowrap');
    td(`<span class="badge ${l.nivelCls}">${esc(l.nivel)}</span>`,'center');
    attachSwipeTranslate(tr, i);
    tb.appendChild(tr);
  });
}
// swipe with translateX
function attachSwipeTranslate(tr, idx){
  let sx=0, sy=0, dx=0, dragging=false;
  function onDown(e){
    const p=e.touches?e.touches[0]:e; sx=p.clientX; sy=p.clientY; dx=0; dragging=true;
    tr.classList.add('dragging'); document.body.style.userSelect='none';
    window.addEventListener('pointermove', onMove); window.addEventListener('pointerup', onUp);
  }
  function onMove(e){
    if(!dragging) return; const p=e.touches?e.touches[0]:e;
    const ndx=p.clientX - sx; const ndy=p.clientY - sy;
    if(Math.abs(ndx) > Math.abs(ndy)){ dx = ndx; tr.style.transform = `translateX(${dx}px)`; }
  }
  function onUp(){
    dragging=false; tr.classList.remove('dragging'); document.body.style.userSelect='auto';
    window.removeEventListener('pointermove', onMove); window.removeEventListener('pointerup', onUp);
    const TH=100;
    if(dx > TH){ openEdit(idx); tr.style.transform='translateX(0)'; }
    else if(dx < -TH){ if(confirm('Excluir este registro?')){ STATE.linhas.splice(idx,1); save(LS_KEYS.linhas,STATE.linhas); renderTable(); } else { tr.style.transform='translateX(0)'; } }
    else { tr.style.transform='translateX(0)'; }
  }
  tr.addEventListener('pointerdown', onDown, {passive:true});
}

// CADASTRO REGISTRO
function mountCadastroRegistro(){
  const catalog=getCatalog();
  document.body.innerHTML=`
  <header class="topbar"><h1>Matriz de Risco — Adicionar registro</h1><div class="actions"><a class="btn" href="./index.html">← Voltar</a></div></header>
  <main class="container"><section class="card"><h2>Preencha os campos</h2>
  <form id="form" class="grid-4">
    <label>Agente <select id="ag"><option value="">Selecione</option></select></label>
    <label>Fator de risco <select id="fa"><option value="">Selecione</option></select></label>
    <label>Possível dano <input id="da"></label>
    <label>Fund. legal <input id="le"></label>
    <label>Fonte geradora <input id="fo" list="fonte-sug"><datalist id="fonte-sug"></datalist></label>
    <label>EPC <div class="inline-field"><select id="epc"></select><button type="button" id="add-epc" class="btn">+ novo</button></div></label>
    <label>EPI <div class="inline-field"><select id="epi"></select><button type="button" id="add-epi" class="btn">+ novo</button></div></label>
    <label>Atenuação <input id="at"></label>
    <label>Intens./Concent. <input id="it"></label>
    <label>Técnica <div class="inline-field"><select id="te"></select><button type="button" id="add-te" class="btn">+ nova</button></div></label>
    <label>Tipo de exposição <div class="inline-field"><select id="ti"></select><button type="button" id="add-ti" class="btn">+ novo</button></div></label>
    <label>Probabilidade <select id="pr"></select></label>
    <label>Severidade <select id="se"></select></label>
    <div class="form-footer"><div id="badge" class="badge">Avaliação do risco: —</div><div class="actions"><button type="reset" class="btn">Limpar</button><button class="btn primary">Salvar registro</button></div></div>
  </form>
  <div id="ok" style="display:none;margin-top:10px;color:#16a34a;font-weight:600">Registro salvo!</div>
  </section></main>`;
  const el={ag:document.getElementById('ag'), fa:document.getElementById('fa'), da:document.getElementById('da'), le:document.getElementById('le'),
    fo:document.getElementById('fo'), fs:document.getElementById('fonte-sug'),
    epc:document.getElementById('epc'), epi:document.getElementById('epi'),
    at:document.getElementById('at'), it:document.getElementById('it'), te:document.getElementById('te'), ti:document.getElementById('ti'),
    pr:document.getElementById('pr'), se:document.getElementById('se'), badge:document.getElementById('badge'), form:document.getElementById('form'), ok:document.getElementById('ok')};
  Object.keys(catalog.agentes).forEach(a=>{const o=document.createElement('option');o.value=a;o.textContent=a; el.ag.appendChild(o)});
  function fill(sel,arr){sel.innerHTML=""; arr.forEach(v=>{const o=document.createElement('option');o.value=v;o.textContent=v; sel.appendChild(o)})}
  function ensure(sel,val){ if(!val) return; if(![...sel.options].some(o=>o.value===val)){const o=document.createElement('option'); o.value=val; o.textContent=val; sel.appendChild(o);} sel.value=val; }
  fill(el.epc, catalog.epc); fill(el.epi, catalog.epi); fill(el.te, catalog.tecnicas); fill(el.ti, catalog.tiposExposicao);
  CATALOG_BASE.probOpts.forEach(o=>{const op=document.createElement('option');op.value=o.value;op.textContent=`${o.value} — ${o.label}`;el.pr.appendChild(op)});
  CATALOG_BASE.sevOpts.forEach(o=>{const op=document.createElement('option');op.value=o.value;op.textContent=`${o.value} — ${o.label}`;el.se.appendChild(op)});
  el.pr.value=2; el.se.value=2; update();
  el.ag.onchange=chgAg; el.fa.onchange=chgFa; el.pr.onchange=update; el.se.onchange=update; el.form.onsubmit=save;
  document.getElementById('add-epc').onclick=()=>addNew('epc', el.epc);
  document.getElementById('add-epi').onclick=()=>addNew('epi', el.epi);
  document.getElementById('add-te').onclick=()=>addNew('tecnicas', el.te);
  document.getElementById('add-ti').onclick=()=>addNew('tiposExposicao', el.ti);
  function addNew(key, sel){ const v=prompt('Digite o novo item:'); if(!v) return; const u=getUser(); if(!u[key]) u[key]=[]; if(!u[key].includes(v)) u[key].push(v); setUser(u); const c=getCatalog(); if(key==='epc') fill(sel,c.epc); if(key==='epi') fill(sel,c.epi); if(key==='tecnicas') fill(sel,c.tecnicas); if(key==='tiposExposicao') fill(sel,c.tiposExposicao); sel.value=v; }
  function chgAg(){ el.fa.innerHTML='<option value="">'+'Selecione'+'</option>'; el.fo.value=''; el.fs.innerHTML=''; (catalog.agentes[el.ag.value]||[]).forEach(f=>{const o=document.createElement('option');o.value=f.factor; o.textContent=f.factor; el.fa.appendChild(o)}); chgFa(); }
  function chgFa(){ const a=el.ag.value, f=el.fa.value; const found=(catalog.agentes[a]||[]).find(x=>x.factor===f);
    if(found){ el.da.value=found.dano||''; el.le.value=found.legal||''; el.fo.value=''; el.fs.innerHTML=''; (found.fontes||[]).forEach(s=>{const o=document.createElement('option'); o.value=s; el.fs.appendChild(o);}); if(found.fontes&&found.fontes.length) el.fo.value=found.fontes[0];
      if(found.defaults){ ensure(el.te,found.defaults.tecnica); ensure(el.ti,found.defaults.tipoExpo); ensure(el.epc,found.defaults.epc); ensure(el.epi,found.defaults.epi); }
    } else { el.da.value=''; el.le.value=''; el.fo.value=''; el.fs.innerHTML=''; }
  }
  function update(){ const r=matrixResult(el.pr.value,el.se.value); el.badge.className='badge '+r.cls; el.badge.textContent='Avaliação do risco: '+r.nivel; }
  function save(ev){ ev.preventDefault(); if(!el.ag.value||!el.fa.value){alert('Selecione agente e fator.');return} const r=matrixResult(el.pr.value,el.se.value);
    const item={agente:el.ag.value,fator:el.fa.value,dano:el.da.value,legal:el.le.value,fonteGeradora:el.fo.value,epc:el.epc.value,epi:el.epi.value,atenuacao:el.at.value,intensidade:el.it.value,tecnica:el.te.value,tipoExpo:el.ti.value,prob:Number(el.pr.value),probLabel:({1:'Possível',2:'Improvável',3:'Pouco provável',4:'Provável'})[el.pr.value],sev:Number(el.se.value),sevLabel:({1:'Leve',2:'Moderado',3:'Sério',4:'Incapacitante'})[el.se.value],nivel:r.nivel,nivelCls:r.cls,createdAt:Date.now()};
    const linhas=load(LS_KEYS.linhas,[]); linhas.unshift(item); save(LS_KEYS.linhas,linhas); el.ok.style.display='block'; setTimeout(()=>el.ok.style.display='none',1500); el.form.reset(); el.pr.value=2; el.se.value=2; update();
  }
}

// CADASTRO FATOR
function mountCadastroFator(){
  const c=getCatalog();
  document.body.innerHTML=`
  <header class="topbar"><h1>Matriz de Risco — Cadastrar novo fator de risco</h1><div class="actions"><a class="btn" href="./index.html">← Voltar</a></div></header>
  <main class="container"><section class="card"><h2>Novo fator</h2>
  <form id="form" class="grid-4">
    <label>Agente <select id="ag"><option value="">Selecione</option></select></label>
    <label>Nome do fator <input id="fa" placeholder="ex.: Poeira com sílica"></label>
    <label>Possível dano <input id="da"></label>
    <label>Fund. legal / Limite <input id="le"></label>
    <label>Fontes geradoras (separe por ;) <input id="fo"></label>
    <label>Técnica sugerida <select id="te"><option value="">(opcional)</option></select></label>
    <label>Tipo de exposição sugerido <select id="ti"><option value="">(opcional)</option></select></label>
    <label>EPC sugerido <select id="epc"><option value="">(opcional)</option></select></label>
    <label>EPI sugerido <select id="epi"><option value="">(opcional)</option></select></label>
    <div class="form-footer"><div class="actions"><button type="reset" class="btn">Limpar</button><button class="btn primary">Salvar fator</button></div></div>
  </form>
  <div id="ok" style="display:none;margin-top:10px;color:#16a34a;font-weight:600">Fator salvo!</div>
  </section>
  <section class="card"><h2>Fatores cadastrados (seus)</h2>
    <table class="factor-list"><thead><tr><th>Agente</th><th>Fator</th><th>Ações</th></tr></thead><tbody id="tb"><tr><td colspan="3">Nenhum ainda.</td></tr></tbody></table>
  </section>
  </main>`;
  const el={ag:document.getElementById('ag'),fa:document.getElementById('fa'),da:document.getElementById('da'),le:document.getElementById('le'),fo:document.getElementById('fo'),
            te:document.getElementById('te'),ti:document.getElementById('ti'),epc:document.getElementById('epc'),epi:document.getElementById('epi'),ok:document.getElementById('ok'),tb:document.getElementById('tb'),form:document.getElementById('form')};
  Object.keys(CATALOG_BASE.agentes).forEach(a=>{const o=document.createElement('option');o.value=a;o.textContent=a; el.ag.appendChild(o)});
  function fill(sel,arr,op){ sel.innerHTML = op? '<option value="">'+'(opcional)'+'</option>' : ''; arr.forEach(v=>{const o=document.createElement('option'); o.value=v; o.textContent=v; sel.appendChild(o);})}
  const cat=getCatalog(); fill(el.te,cat.tecnicas,true); fill(el.ti,cat.tiposExposicao,true); fill(el.epc,cat.epc,true); fill(el.epi,cat.epi,true);
  el.form.onsubmit = (e)=>{ e.preventDefault(); if(!el.ag.value||!el.fa.value.trim()){alert('Agente e nome do fator são obrigatórios.');return} const novo={factor:el.fa.value.trim(),dano:el.da.value.trim(),legal:el.le.value.trim(),fontes:el.fo.value.split(';').map(s=>s.trim()).filter(Boolean),defaults:{}}; if(el.te.value) novo.defaults.tecnica=el.te.value; if(el.ti.value) novo.defaults.tipoExpo=el.ti.value; if(el.epc.value) novo.defaults.epc=el.epc.value; if(el.epi.value) novo.defaults.epi=el.epi.value; const u=getUser(); if(!u.agentes[el.ag.value]) u.agentes[el.ag.value]=[]; u.agentes[el.ag.value].push(novo); setUser(u); el.ok.style.display='block'; setTimeout(()=>el.ok.style.display='none',1500); el.form.reset(); renderList(); };
  renderList();
  function renderList(){ const u=getUser(); const list=[]; Object.keys(u.agentes).forEach(a=> (u.agentes[a]||[]).forEach((f,ix)=> list.push({a,f,ix}) )); el.tb.innerHTML=""; if(!list.length){ const tr=document.createElement('tr'); const td=document.createElement('td'); td.colSpan=3; td.textContent="Nenhum ainda."; tr.appendChild(td); el.tb.appendChild(tr); return; } list.forEach(it=>{ const tr=document.createElement('tr'); const tda=document.createElement('td'); tda.textContent=it.a; tr.appendChild(tda); const tdf=document.createElement('td'); tdf.textContent=it.f.factor; tr.appendChild(tdf); const tdx=document.createElement('td'); const b=document.createElement('button'); b.className='btn danger'; b.textContent='Excluir'; b.onclick=()=>{ if(confirm('Excluir este fator?')){ const u=getUser(); u.agentes[it.a].splice(it.ix,1); if(!u.agentes[it.a].length) delete u.agentes[it.a]; setUser(u); renderList(); } }; tdx.appendChild(b); tr.appendChild(tdx); el.tb.appendChild(tr); }); }
}

// XLSX export based on template
function exportXLSX(){
  const rows = STATE.linhas;
  const xml = (s)=> String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const cInline = (v, s=null, r=null)=> `<c t="inlineStr" ${r?`r="${r}"`:""} ${s!=null?`s="${s}"`:""}><is><t xml:space="preserve">${xml(v||"")}</t></is></c>`;
  const STYLE = {header:8, body:3, risco:{"Risco Irrelevante":8, "Risco Baixo":6, "Risco Médio":5, "Risco Alto":7, "Risco Crítico":7}};
  const headers = ["Agente","Fator de Risco","Possível dano","Fund. legal / Limite de exposição","Fonte geradora","EPC","EPI","Atenuação","Intens./Concent.","Técnica","Tipo de exposição","Probabilidade","Severidade","Nível de risco"];
  const headerRow = `<row r="1">` + headers.map((h,i)=> cInline(h, i===8?9:STYLE.header, String.fromCharCode(65+i)+"1")).join("") + `</row>`;
  const dataRows = rows.map((l,idx)=>{
    const r = matrixResult(l.prob,l.sev);
    const sRisk = STYLE.risco[r.nivel] ?? STYLE.body;
    const Rn = 2+idx;
    const cols = [
      cInline(l.agente, STYLE.body, "A"+Rn),
      cInline(l.fator, STYLE.body, "B"+Rn),
      cInline(l.dano, STYLE.body, "C"+Rn),
      cInline(l.legal, STYLE.body, "D"+Rn),
      cInline(l.fonteGeradora, STYLE.body, "E"+Rn),
      cInline(l.epc, STYLE.body, "F"+Rn),
      cInline(l.epi, STYLE.body, "G"+Rn),
      cInline(l.atenuacao, STYLE.body, "H"+Rn),
      cInline(l.intensidade, STYLE.body, "I"+Rn),
      cInline(l.tecnica, STYLE.body, "J"+Rn),
      cInline(l.tipoExpo, STYLE.body, "K"+Rn),
      cInline(`${l.prob} — ${l.probLabel||""}`, 4, "L"+Rn),
      cInline(`${l.sev} — ${l.sevLabel||""}`, 4, "M"+Rn),
      cInline(r.nivel, sRisk, "N"+Rn),
    ];
    return `<row r="${Rn}">`+cols.join("")+`</row>`;
  }).join("");
  const last = Math.max(1, rows.length+1);
  const dim = `A1:N${last}`;
  const sheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dim}"/>
  ${TPL.sheetViews}
  <sheetFormatPr defaultRowHeight="16"/>
  ${TPL.cols}
  <sheetData>
    ${headerRow}
    ${dataRows}
  </sheetData>
  <autoFilter ref="A1:N${last}"/>
</worksheet>`;
  const files = {
    "[Content_Types].xml": TPL.contentTypes,
    "_rels/.rels": TPL.rels,
    "xl/workbook.xml": TPL.workbook,
    "xl/_rels/workbook.xml.rels": TPL.workbookRels,
    "xl/styles.xml": TPL.styles,
    "xl/theme/theme1.xml": TPL.theme1,
    "xl/sharedStrings.xml": TPL.sharedStrings,
    "xl/worksheets/sheet1.xml": sheet,
  };
  const blob = buildZip(files);
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='matriz_riscos.xlsx'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),2000);
}
// ZIP (store)
function buildZip(obj){
  const enc=new TextEncoder(); const files=Object.keys(obj).map(n=>({name:n,data:enc.encode(obj[n])}));
  const cd=[]; const out=[]; let off=0;
  files.forEach(f=>{ const crc=crc32(f.data); const nm=new TextEncoder().encode(f.name);
    out.push(u32(0x04034b50),u16(20),u16(0),u16(0),u16(0),u16(0),u32(crc),u32(f.data.length),u32(f.data.length),u16(nm.length),u16(0)); out.push(nm,f.data);
    cd.push(u32(0x02014b50),u16(0x031E),u16(20),u16(0),u16(0),u16(0),u16(0),u32(crc),u32(f.data.length),u32(f.data.length),u16(nm.length),u16(0),u16(0),u16(0),u16(0),u32(0),u32(off),nm);
    off += 30 + nm.length + f.data.length;
  });
  const cdBytes=concat(cd); const outBytes=concat(out);
  const end=concat([u32(0x06054b50),u16(0),u16(0),u16(files.length),u16(files.length),u32(cdBytes.length),u32(outBytes.length),u16(0)]);
  return new Blob([outBytes,cdBytes,end],{type:"application/zip"});
  function u16(n){const b=new Uint8Array(2);b[0]=n&255;b[1]=(n>>>8)&255;return b}
  function u32(n){const b=new Uint8Array(4);b[0]=n&255;b[1]=(n>>>8)&255;b[2]=(n>>>16)&255;b[3]=(n>>>24)&255;return b}
  function concat(arr){let len=0;arr.forEach(a=>len+=a.length);const r=new Uint8Array(len);let o=0;arr.forEach(a=>{r.set(a,o);o+=a.length});return r}
  function crc32(buf){let c=~0;for(let i=0;i<buf.length;i++){c=(c>>>8)^CRC_TABLE[(c^buf[i])&0xFF]}return ~c>>>0}
}
const CRC_TABLE=(()=>{let c,t=[];for(let n=0;n<256;n++){c=n;for(let k=0;k<8;k++){c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1)}t[n]=c>>>0}return t})();
// boot
(function boot(){ if(document.getElementById('page-index')) mountIndex(); else if(document.getElementById('page-cadastro-registro')) mountCadastroRegistro(); else if(document.getElementById('page-cadastro-fator')) mountCadastroFator(); else mountIndex(); })();
