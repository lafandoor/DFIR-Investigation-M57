const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat, BorderStyle,
  WidthType, ShadingType, SimpleField, PageBreak
} = require('docx');
const fs = require('fs');


const C = {
  navy:    "1B3A5C", steel:   "2E6DA4", crimson: "C0392B",
  amber:   "D4850A", green:   "1E7145", tblHead: "1B3A5C",
  tblAlt:  "EAF0F8", tblBdr:  "B0C4DE", white:   "FFFFFF",
  charcoal:"2C2C2C", gray:    "6C6C6C",
};

const bdr  = (c=C.tblBdr) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const bdrs = (c=C.tblBdr) => ({ top:bdr(c), bottom:bdr(c), left:bdr(c), right:bdr(c) });

const run    = (t,o={}) => new TextRun({text:t, font:"Arial", size:20, color:C.charcoal, ...o});
const bold   = (t,o={}) => run(t,{bold:true,...o});
const italic = (t,o={}) => run(t,{italics:true,...o});
const mono   = (t,o={}) => new TextRun({text:t, font:"Courier New", size:18, color:C.navy, ...o});
const mbold  = (t,o={}) => new TextRun({text:t, font:"Courier New", size:18, color:C.crimson, bold:true, ...o});

const h1 = t => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children:[new TextRun({text:t, font:"Arial", size:36, bold:true, color:C.navy})],
  spacing:{before:360,after:120},
  border:{bottom:{style:BorderStyle.SINGLE,size:8,color:C.steel,space:4}}
});
const h2 = t => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children:[new TextRun({text:t, font:"Arial", size:28, bold:true, color:C.steel})],
  spacing:{before:280,after:100}
});
const h3 = t => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children:[new TextRun({text:t, font:"Arial", size:24, bold:true, color:C.navy})],
  spacing:{before:200,after:80}
});
const para = (ch,o={}) => new Paragraph({
  children: Array.isArray(ch)?ch:[run(ch)],
  spacing:{after:120}, ...o
});
const bullet = ch => new Paragraph({
  numbering:{reference:"bullets",level:0},
  children:Array.isArray(ch)?ch:[run(ch)],
  spacing:{after:60}
});
const numbered = ch => new Paragraph({
  numbering:{reference:"numbers",level:0},
  children:Array.isArray(ch)?ch:[run(ch)],
  spacing:{after:60}
});
const sp = (n=80) => new Paragraph({children:[run("")],spacing:{after:n}});
const pb = () => new Paragraph({children:[new PageBreak()]});

const alertBox = (label, txt, col=C.crimson) => new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders:{top:{style:BorderStyle.SINGLE,size:12,color:col},
             bottom:{style:BorderStyle.SINGLE,size:4,color:col},
             left:{style:BorderStyle.SINGLE,size:12,color:col},
             right:{style:BorderStyle.SINGLE,size:4,color:col}},
    shading:{fill:"FAFAFA",type:ShadingType.CLEAR},
    margins:{top:100,bottom:100,left:140,right:140},
    width:{size:9360,type:WidthType.DXA},
    children:[new Paragraph({children:[
      bold(`[${label}]  `,{color:col}),
      run(txt,{color:C.charcoal})
    ],spacing:{after:0}})]
  })]})],
});

const code = lines => new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders:bdrs(C.navy),
    shading:{fill:"1A1A2E",type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:140,right:140},
    width:{size:9360,type:WidthType.DXA},
    children:lines.map(l=>new Paragraph({
      children:[new TextRun({text:l,font:"Courier New",size:17,color:"A8FF78"})],
      spacing:{after:20}
    }))
  })]})],
});

const tbl = (headers, rows, widths) => {
  const total = widths.reduce((a,b)=>a+b,0);
  return new Table({
    width:{size:total,type:WidthType.DXA}, columnWidths:widths,
    rows:[
      new TableRow({tableHeader:true, children:headers.map((h,i)=>new TableCell({
        borders:bdrs(), shading:{fill:C.tblHead,type:ShadingType.CLEAR},
        width:{size:widths[i],type:WidthType.DXA},
        margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[bold(h,{color:C.white})]})]
      }))}),
      ...rows.map((row,ri)=>new TableRow({children:row.map((cell,ci)=>new TableCell({
        borders:bdrs(),
        shading:{fill:ri%2===0?C.white:C.tblAlt,type:ShadingType.CLEAR},
        width:{size:widths[ci],type:WidthType.DXA},
        margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:Array.isArray(cell)?cell:[run(cell)]})]
      }))}))
    ]
  });
};




const ch = [];




ch.push(
  sp(600),
  new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:40},
    children:[bold("CONFIDENTIAL — ACADEMIC SUBMISSION",{size:20,color:C.crimson})]}),
  new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:60},
    border:{bottom:{style:BorderStyle.DOUBLE,size:6,color:C.steel,space:6}},
    children:[new TextRun({text:"DIGITAL FORENSICS & INCIDENT RESPONSE",font:"Arial",size:52,bold:true,color:C.navy})]}),
  sp(60),
  new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:40},
    children:[new TextRun({text:"FORENSIC INVESTIGATION REPORT — FINAL EDITION",font:"Arial",size:34,bold:true,color:C.steel})]}),
  new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:10},
    children:[new TextRun({text:"Unknown Breach: Evidence Discovery, Validation,",font:"Arial",size:28,color:C.charcoal,italics:true})]}),
  new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:200},
    children:[new TextRun({text:"and Full Forensic Investigation",font:"Arial",size:28,color:C.charcoal,italics:true})]}),
  sp(120),
  tbl(["Field","Value"],[
    ["Case Title",         "M57.Biz — Corporate Data Exfiltration Incident"],
    ["Case Number",        "DFIR-2024-012"],
    ["Evidence Image",     "jo-2009-11-16.E01"],
    ["Computer Name",      "M57-JO  (Windows XP SP3, x86)"],
    ["Subject",            "Jo — Employee, M57.biz"],
    ["Examiner",           "Youssef Moataz — AASTMT Cybersecurity, Class of 2026"],
    ["Course",             "Digital Forensics & Incident Response (CCY3102)"],
    ["Institution",        "Arab Academy for Science, Technology & Maritime Transport"],
    ["Date Submitted",     "April 2026"],
    ["Primary Tools",      "Autopsy 4.23  |  FTK Imager 4.7.3.81  |  Velociraptor (simulated)"],
    ["Verified MD5",       "f3160a776ca1d59172c6acb622b67459  ✓  MATCH"],
    ["Verified SHA1",      "cf2d8ca1e282501dc7977acd4ac48df1a4b23e9f  ✓  MATCH"],
    ["Report Version",     "v2.0 — Final Edition (all artefact layers integrated)"],
  ],[2800,6560]),
  sp(120),
  new Paragraph({alignment:AlignmentType.CENTER,
    children:[italic("This report is prepared solely for academic evaluation purposes.",{color:C.gray,size:18})]}),
  pb(),
);




ch.push(
  h1("Executive Summary"),
  para([
    run("This final edition forensic investigation was conducted against "),
    bold("jo-2009-11-16.E01"),
    run(", sourced from the M57-Patents dataset (Digital Corpora, UCSD). The machine "),
    bold("M57-JO"),
    run(" ran Windows XP SP3 and was registered to employee "),
    bold("Jo"),
    run(" of "),bold("M57.biz"),run(". The image was captured on 16–17 November 2009.")
  ]),
  para([
    run("This edition incorporates analysis of "),
    bold("Event Logs (Application + System), Registry hives (SYSTEM, SOFTWARE, NTUSER.DAT), Prefetch files (86 entries including unallocated residue), keyword search results, web form autofill data, remote-monitoring artefacts,"),
    run(" and metadata. A critical new finding — "),
    bold("ZIPPER.EXE in unallocated Prefetch space"),
    run(" — elevates the exfiltration assessment significantly.")
  ]),
  sp(40),
  alertBox("CRITICAL FINDING — UPGRADED",
    "ZIPPER.EXE (file compression tool) was identified in unallocated Prefetch space, indicating it was previously executed and the Prefetch entry was subsequently overwritten — a hallmark of deliberate cleanup. Combined with the LaCie HDD connection, IMAPI CD-burning activity confirmed via System Event Log (Event 7036 — 2009-11-17 02:31:53), and HELIX.EXE execution immediately followed by an Event Log shutdown (6006 at 05:08:26), the evidence profile is now consistent with premeditated data compression, exfiltration, and partial anti-forensic cleanup.",
    C.crimson),
  sp(100),
  tbl(["Finding","Severity","Evidence Source","Status"],[
    ["ZIPPER.EXE in unallocated Prefetch — file compression tool previously run","CRITICAL","Prefetch_20260429123907.csv","CONFIRMED"],
    ["LaCie Rugged HDD connected at 02:20:07 on acquisition day","CRITICAL","USB_Device_Attached CSV","CONFIRMED"],
    ["IMAPI CD-Burning Service: RUNNING → STOPPED on 2009-11-17 02:31:53–59 (Event 7036)","CRITICAL","sys-event.xlsx","CONFIRMED"],
    ["Fastfat Warning (Event 50) on \\Device\\HarddiskVolume3 at 02:25:28 — removable volume","HIGH","sys-event.xlsx","CONFIRMED"],
    ["HELIX.EXE run at 05:07:23 → EventLog stopped at 05:08:26 (1 min later)","HIGH","Prefetch + sys-event.xlsx","CONFIRMED"],
    ["MDD_1.3.EXE (memory dump) executed at 02:22:34 on acquisition day","HIGH","Prefetch CSV","CONFIRMED"],
    ["webmail.m57.biz: login_username='jo' submitted 5 times — failed login confirmed","HIGH","Web_Form_Autofill CSV","CONFIRMED"],
    ["USPTO.gov patent office browsed extensively (16 pages) before imaging","MEDIUM","Web_History CSV","CONFIRMED"],
    ["mstsc.exe (Remote Desktop Client) flagged as notable — possible lateral movement","MEDIUM","Remote_Monitoring CSV","CONFIRMED"],
    ["DEFRAG.EXE executed 2009-11-15 — possible anti-forensic activity","MEDIUM","Prefetch CSV","CONFIRMED"],
    ["No malicious persistence in Run keys — insider threat confirmed","INFO","Registry / Autopsy","CONFIRMED"],
    ["Keyword search: no literal/regex hits — no plaintext exfil strings in allocated space","INFO","TSK_KEYWORD_HIT CSV","CONFIRMED"],
  ],[3400,1400,2200,1560]),
  sp(100),
  para([
    bold("Exfiltration Assessment: "),
    run("Based on the totality of evidence — ZIPPER.EXE execution (now deleted), LaCie HDD connection, IMAPI CD-burning confirmed by System Event Log, and HELIX.EXE counter-forensic activity — data exfiltration is assessed as "),
    bold("HIGHLY PROBABLE",{color:C.crimson}),
    run(" with medium-high confidence. The specific channel (USB, optical disc, or webmail) cannot be conclusively determined from disk evidence alone.")
  ]),
  pb(),
);




ch.push(
  h1("Step 0 — Evidence Sourcing & Dataset Justification"),
  h2("0.1  Dataset Identification"),
  tbl(["Attribute","Details"],[
    ["Dataset Name",     "M57-Patents Scenario"],
    ["Repository",       "Digital Corpora (digitalcorpora.org) — University of California San Diego"],
    ["Primary Image",    "jo-2009-11-16.E01  (Jo's workstation)"],
    ["Image Format",     "Expert Witness Format (E01) — EnCase-compatible multi-segment"],
    ["OS in Image",      "Microsoft Windows XP Service Pack 3  (x86)"],
    ["Computer Name",    "M57-JO"],
    ["Registered Owner", "Jo  /  Organisation: M57.biz  /  Product ID: 76487-027-5250835-22087"],
    ["Acquisition Date", "16–17 November 2009"],
    ["Image Path Used",  "G:\\projj\\jo-2009-11-16.E01"],
    ["Autopsy Case",     "DFIR_Project  (jo-2009-11-16.E01_1 Host)"],
    ["License",          "Creative Commons — open for academic and research use"],
  ],[2800,6560]),
  sp(100),
  h2("0.2  Justification for Selection"),
  bullet([bold("Realism: "),run("Created by Garfinkel et al. (UCSD) to simulate genuine corporate activity — real user browsing, email, and file system operations, not synthetic artefacts.")]),
  bullet([bold("Scenario Alignment: "),run("The M57 scenario depicts a corporate environment with a suspected insider data leak, directly matching the project brief.")]),
  bullet([bold("Artefact Richness: "),run("Contains Prefetch files (86 entries), Registry hives (SYSTEM, SOFTWARE, SAM, SECURITY, NTUSER.DAT), Event Logs (AppEvent.Evt, SysEvent.Evt, SecEvent.Evt), browser history, email client data, USB entries — all 9 required analysis domains covered.")]),
  bullet([bold("Tool Compatibility: "),run("E01 format natively supported by FTK Imager 4.7.3.81 and Autopsy 4.23 with no conversion.")]),
  bullet([bold("Verified Provenance: "),run("Hosted by Digital Corpora (UCSD) with documented chain of custody. Cited in peer-reviewed DFIR literature and used in university curricula worldwide.")]),
  sp(100),
  h2("0.3  Dataset Limitations"),
  bullet([bold("LaCie Drive Not Imaged: "),run("The LaCie Rugged HDD connected before acquisition is the most critical physical evidence item but was not included in the E01 image. Its content remains unknown from this investigation alone.")]),
  bullet([bold("No RAM Image Available for Analysis: "),run("Although MDD_1.3.EXE was run and a /RAM directory is referenced in Prefetch, the resulting memory dump is not part of the E01. Volatile state (running processes, network connections, decryption keys) is unavailable.")]),
  bullet([bold("No Network Capture: "),run("PCAP/NetFlow data is unavailable. Web-based upload channels (HTTPS, webmail) cannot be confirmed or denied from disk alone.")]),
  bullet([bold("ShimCache Not Parsed: "),run("The SYSTEM hive was confirmed present (3.4 MB, last modified 2009-11-17 02:22:31 EET). ShimCache parsing was not completed due to tooling limitations — see Task 4.5.")]),
  bullet([bold("Amcache Not Applicable: "),run("Windows XP SP3 predates the Amcache artefact (introduced Windows 8). Not applicable — see Task 4.6.")]),
  bullet([bold("Security Event Log Sparse: "),run("SecEvent.Evt was captured but contained no entries post-2009-11-08 — it was cleared or overwritten during initial setup, providing limited logon/object access data.")]),
  bullet([bold("2009 XP SP3 Context: "),run("Artefact locations differ from modern Windows. Notable absences include Jump Lists (Win7+) and VSS Shadow Copies (not detected).")]),
  pb(),
);




ch.push(
  h1("Task 1 — Evidence Validation"),
  h2("1.1  Tool: FTK Imager 4.7.3.81"),
  para("The E01 image was loaded into FTK Imager (Exterro v4.7.3.81). FTK reads acquisition hashes from E01 case metadata and recomputes them sector-by-sector for verification."),
  h2("1.2  Hash Verification Results"),
  tbl(["Algorithm","Stored Hash (Acquisition)","Computed Hash (Verification)","Result"],[
    ["MD5", "f3160a776ca1d59172c6acb622b67459","f3160a776ca1d59172c6acb622b67459","✓  MATCH"],
    ["SHA1","cf2d8ca1e282501dc7977acd4ac48df1a4b23e9f","cf2d8ca1e282501dc7977acd4ac48df1a4b23e9f","✓  MATCH"],
  ],[1400,3200,3200,1560]),
  sp(80),
  code([
    "FTK Imager v4.7.3.81 — Drive/Image Verify Results",
    "──────────────────────────────────────────────────────────────",
    "Name              : jo-2009-11-16.E01",
    "Sector count      : 25,429,824",
    "MD5  → Computed   : f3160a776ca1d59172c6acb622b67459   [ MATCH ✓ ]",
    "SHA1 → Computed   : cf2d8ca1e282501dc7977acd4ac48df1a4b23e9f",
    "Bad blocks        : No bad blocks found in image  [ CLEAN ✓ ]",
    "INTEGRITY RESULT  : VERIFIED — Image is forensically sound.",
  ]),
  sp(80),
  h2("1.3  Registry Hive Identification"),
  para([run("Registry hives confirmed present in "),mono("C:\\WINDOWS\\system32\\config\\"),run(" (source: config_20260428115929.csv):")]),
  tbl(["Hive File","Size","Last Modified (EET)","Contents"],[
    ["system",    "3.4 MB", "2009-11-17 02:22:31","SYSTEM hive — USBSTOR, services, ShimCache (AppCompatCache)"],
    ["software",  "11.5 MB","2009-11-17 02:32:38","SOFTWARE hive — installed programs, Run keys"],
    ["SAM",       "256 KB", "2009-11-17 02:32:38","SAM — user account database"],
    ["SECURITY",  "256 KB", "2009-11-17 02:32:38","SECURITY — policy, audit settings"],
    ["AppEvent.Evt","64 KB","2009-11-17 02:32:16","Application Event Log — 91 events"],
    ["SysEvent.Evt","128 KB","2009-11-17 02:32:16","System Event Log — 169 events"],
    ["SecEvent.Evt","64 KB","2009-11-08 17:53:06","Security Event Log — minimal (cleared during setup)"],
    ["NTUSER.DAT", "N/A",  "N/A",                 "Jo's user hive — UserAssist, shell bags, MRU"],
  ],[1800,1200,2200,4160]),
  sp(80),
  alertBox("FINDING 1.1","Hash verification confirms the forensic image is intact and unmodified since acquisition. All registry hives and event logs are present and accessible. Chain of custody is established.", C.green),
  pb(),
);




ch.push(
  h1("Task 2 — Initial Triage"),
  h2("2.1  Operating System Identification"),
  tbl(["Parameter","Value"],[
    ["OS Name",            "Microsoft Windows XP Service Pack 3"],
    ["Architecture",       "x86  (32-bit)"],
    ["Computer Name",      "M57-JO"],
    ["Registered Owner",   "Jo"],
    ["Organisation",       "M57.biz"],
    ["Windows Path",       "C:\\WINDOWS"],
    ["Temp Directory",     "%SystemRoot%\\TEMP"],
    ["Product ID",         "76487-027-5250835-22087"],
    ["First Boot",         "2009-11-08 17:38:17 EET  (EventLog first started — MACHINENAME)"],
    ["Last EventLog Stop", "2009-11-17 02:32:15 EET  (Event 6006 — system shutdown/imaging)"],
  ],[3000,6360]),
  sp(100),
  h2("2.2  User Accounts"),
  tbl(["Username","Type","Profile Path","Notes"],[
    ["Jo",            "Standard User","C:\\Documents and Settings\\Jo\\",            "Primary subject — all user-level artefacts"],
    ["Administrator", "Administrator","C:\\Documents and Settings\\Administrator\\",  "Used for system setup: AVG, Java, OpenOffice installs"],
  ],[1600,1400,3200,3160]),
  sp(100),
  h2("2.3  Key Installed Software (Chronological)"),
  tbl(["Program","Install Date (EET)","Forensic Relevance"],[
    ["AVG 9.0",              "2009-11-09 01:14:58","Antivirus — active throughout; AV logs may contain file scan history"],
    ["OpenOffice.org 3.1",   "2009-11-10 01:01:42","Document editor — Definitions.odt recently accessed"],
    ["Java 6 Update 17",     "2009-11-10 03:31:48","Confirmed installed via AppEvent 11707 — JRE for Python/OO"],
    ["Mozilla Firefox 3.5.5","2009-11-13 01:52:06","Primary browser — installed on day of HELIX activity"],
    ["Python 2.6.4",         "2009-11-13 03:59:27","Confirmed via AppEvent 11707 — scripting runtime"],
    ["WebFldrs XP",          "2009-11-11 02:12:54","Confirmed via AppEvent 11728 — web folder client"],
    ["mstsc.exe",            "Built-in XP",         "Remote Desktop Client — flagged as notable; accessed 2009-11-11"],
  ],[2600,2200,4560]),
  sp(100),
  h2("2.4  Autopsy Artefact Summary"),
  tbl(["Category","Count","Priority"],[
    ["Web History (Firefox + IE)",  "3,557","🔴 HIGH"],
    ["Web Bookmarks",               "96",   "🟡 MEDIUM"],
    ["Web Cookies",                 "309",  "⚪ LOW"],
    ["Web Search",                  "36",   "🔴 HIGH"],
    ["Web Downloads",               "3",    "🔴 HIGH"],
    ["Web Form Autofill",           "8",    "🔴 HIGH — login_username='jo' × 5"],
    ["USB Device Attached",         "6",    "🔴 CRITICAL"],
    ["E-Mail Messages",             "2",    "🟡 MEDIUM — old test msgs only"],
    ["Run Programs (Prefetch)",     "49",   "🔴 HIGH"],
    ["Prefetch Files (raw)",        "86",   "🔴 HIGH — includes unallocated entries"],
    ["Shell Bags",                  "25",   "🟡 MEDIUM"],
    ["Recent Documents",            "2",    "🟡 MEDIUM"],
    ["Installed Programs",          "95",   "🟡 MEDIUM"],
    ["Extension Mismatch",          "94",   "🟡 MEDIUM"],
    ["Suspicious Items (YARA)",     "1,350","⚪ LOW — primarily OO/Java FP"],
    ["Keyword Hits (literal/regex)","0",    "INFO — no plaintext IOC strings found"],
    ["Email Address Hits",          "3,515","INFO — regex pattern across all files"],
  ],[3000,1200,5160]),
  pb(),
);




ch.push(
  h1("Task 3 — Deep Disk Forensics"),
  h2("3.1  Browser History — Critical Findings"),
  h3("3.1.1  M57.biz Internal Webmail Access"),
  tbl(["Timestamp (EET)","URL / Page Title","Significance"],[
    ["2009-11-16 20:48:23","webmail.m57.biz/src/login.php — SquirrelMail Login","First webmail session"],
    ["2009-11-16 20:48:38","mailboxes.m57.biz — Mailbox Manager 1.0","Mail account manager accessed"],
    ["2009-11-16 21:33:24","MSIMN.EXE (Outlook Express) runs","Local email client opened simultaneously"],
    ["2009-11-16 21:34:13","webmail.m57.biz/src/login.php","Second webmail session"],
    ["2009-11-16 21:35:40","webmail.m57.biz/src/redirect.php — 'Unknown user or password incorrect.'","⚠ FAILED LOGIN"],
    ["2009-11-16 21:36:06–24","Three consecutive login retries","⚠ REPEATED CREDENTIAL FAILURE"],
    ["2009-11-16 21:36:24","mailboxes.m57.biz — Mailbox Manager","Access gained after retries"],
    ["2009-11-16 22:56:15","mailboxes.m57.biz/?restart=1","Further mailbox session"],
  ],[2200,4000,3160]),
  sp(80),
  alertBox("FINDING 3.1 — NEW CONFIRMATION",
    "Web Form Autofill data (formhistory.sqlite) confirms that the username 'jo' was auto-submitted 5 times to the webmail login form (created 2009-11-16 20:48:58, last accessed 2009-11-16 21:36:43). This proves Jo was the specific user operating the browser during all webmail sessions — not a third party using the machine.",
    C.crimson),
  sp(100),
  h3("3.1.2  USPTO.gov — Patent Office Research"),
  tbl(["Timestamp (EET)","URL","Notes"],[
    ["2009-11-16 21:03:44","www.uspto.gov","First visit — general navigation"],
    ["2009-11-17 02:01:31","www.uspto.gov","Second extensive session begins"],
    ["2009-11-17 02:01:39","uspto.gov/patents/basics/index.html","General Information Concerning Patents"],
    ["2009-11-17 02:01:57","uspto.gov/patents/publications/index.jsp","Office of Data Management"],
    ["2009-11-17 02:02:06","uspto.gov/products/index.jsp","Products & Services"],
    ["2009-11-17 02:02:15","uspto.gov/patents/process/search/","Search for Patents"],
    ["2009-11-17 02:02:26","uspto.gov/inventors/index.jsp","Inventors Resources"],
    ["2009-11-17 02:09:12","uspto.gov/patents/index.jsp","Patents main index"],
    ["2009-11-17 02:09:24","uspto.gov/patents/resources/types/index.jsp","Patents Guidance, Tools & Manuals"],
  ],[2200,3800,3360]),
  sp(100),
  h2("3.2  Web Search Analysis (36 Total Searches)"),
  para("Keyword searches conducted against all evidence (TSK_KEYWORD_HIT CSV) returned zero hits for both Single Literal Keyword Search and Single Regular Expression Search. The Email Addresses regex matched 3,515 files — all representing email address strings embedded across web cache, mbox files, and browser databases. No targeted exfiltration-related search strings (e.g., 'upload', 'encrypt', 'send file', 'delete history') were found in allocated disk space."),
  sp(40),
  alertBox("FINDING 3.2 — KEYWORD SEARCH INTERPRETATION",
    "The zero literal/regex keyword hits indicate that: (1) Jo did not save documents containing exfiltration-planning language in plaintext, (2) suspicious content was either not written to disk or was successfully deleted, and (3) the absence of keyword hits is consistent with a user who was aware of forensic investigation techniques — corroborated by HELIX.EXE execution. The 3,515 email address pattern matches are artefacts of embedded email headers in browser cache, mbox files, and OpenOffice documents — not indicative of standalone exfiltration.",
    C.amber),
  sp(100),
  h2("3.3  Web Downloads"),
  tbl(["File","URL / Source","Date (EET)","Risk"],[
    ["python-2.6.4.msi","python.org/ftp/python/2.6.4/","2009-11-13 03:55:49","LOW — legitimate; confirmed installed (AppEvent 11707)"],
    ["Firefox Setup 3.5.5.exe","Zone.Identifier present","2009-11-13","LOW — legitimate browser setup"],
    ["R79733[1].EXE","IE Temp Files — Administrator account","Unknown","LOW — likely AVG update package"],
  ],[2400,3200,1800,2160]),
  sp(100),
  h2("3.4  Recent Documents"),
  tbl(["File","Path","Last Accessed (EET)","Notes"],[
    ["Definitions.odt","C:\\Documents and Settings\\Jo\\My Documents\\","2009-11-17 02:10:09","Accessed 10 min before LaCie HDD connection — content unknown"],
    ["devmgmt.msc","C:\\WINDOWS\\system32\\","Not dated","Device Manager — may relate to USB device management"],
  ],[2200,3200,2000,2160]),
  sp(100),
  h2("3.5  Deleted Files — Keyword Search Results"),
  para("No plaintext exfiltration artefacts were recovered from deleted file space via keyword search. The 94 Extension Mismatch items are system components with incorrect extensions, not user-created mismatched files. The 1,350 Suspicious Items are dominated by OpenOffice gallery files (.sdv), Java cache entries, and Windows cookie index files — all legitimate system artefacts generating YARA false positives."),
  sp(100),
  h2("3.6  Remote Monitoring / Management Tools Detected"),
  tbl(["File","Path","Score","Category","Notes"],[
    ["log.txt (Jo)","...Jo\\AppData\\OpenOffice.org\\...\\uno_packages\\cache\\log.txt","Likely Notable","Atera","2009-11-16 20:47:36 — coincides with WMIC/JAVA activity"],
    ["log.txt (Admin)","...Administrator\\AppData\\OpenOffice.org\\...\\cache\\log.txt","Likely Notable","Atera","2009-11-10 03:24:34 — during Java install"],
    ["system.LOG","C:\\WINDOWS\\system32\\config\\system.LOG","Likely Notable","Kaseya (VSA)","YARA match on SYSTEM hive log — likely false positive"],
    ["mstsc.exe","C:\\WINDOWS\\system32\\mstsc.exe","Likely Notable","Remote Desktop (mstsc)","Accessed 2009-11-11 22:02:10 — Jo may have RDP'd to another machine"],
  ],[1800,3200,1400,1800,1360]),
  sp(80),
  alertBox("FINDING 3.3 — REMOTE DESKTOP",
    "mstsc.exe (Windows Remote Desktop Client) was last accessed on 2009-11-11 22:02:10 EET and was flagged as a Remote Monitoring/Management tool. If Jo used RDP to connect to another internal M57.biz machine, data could have been exfiltrated laterally — a channel not visible in this single-endpoint image. Server-side RDP event logs should be examined.",
    C.amber),
  pb(),
);




ch.push(
  h1("Task 4 — Windows Artifact Analysis"),
  h2("4.1  Event Log Analysis"),
  h3("4.1.1  Application Event Log (AppEvent.Evt — 91 events)"),
  para([run("Source: "),mono("C:\\WINDOWS\\system32\\config\\AppEvent.Evt"),run(" — Last modified 2009-11-17 02:32:16 EET")]),
  tbl(["Event ID","Timestamp (EET)","Source","Description","Significance"],[
    ["11707","2009-11-13 03:59:27","MsiInstaller","Python 2.6.4 — Installation completed successfully","Confirms Python install timing"],
    ["11707","2009-11-10 03:31:48","MsiInstaller","Java 6 Update 17 — Installation operation completed","Confirms Java install"],
    ["11728","2009-11-11 02:12:54","MsiInstaller","WebFldrs XP — Configuration completed","WebDAV folders configured"],
    ["1800", "2009-11-13 06:49:43","SecurityCenter","Security Center service started","Normal boot sequence"],
    ["1",    "2009-11-13 06:49:45","avg9emc",      "AVG Email Scanner: Service started","AV active — email scanning"],
    ["1",    "2009-11-12 19:16:46","avg9emc",      "AVG Email Scanner: Service stopped","AV restart cycle"],
    ["1",    "2009-11-12 19:17:22","avg9emc",      "AVG Email Scanner: Service started","AV restart"],
    ["7",    "2009-11-13 04:23:31","crypt32",      "Certificate store update from windowsupdate.com","Certificate refresh — normal"],
    ["2",    "2009-11-13 04:23:32","crypt32",      "Certificate download from windowsupdate.com","Normal Windows Update activity"],
  ],[1100,2200,1600,3000,1460]),
  sp(100),
  h3("4.1.2  System Event Log (SysEvent.Evt — 169 events) — Forensically Critical"),
  para([run("Source: "),mono("C:\\WINDOWS\\system32\\config\\SysEvent.Evt"),run(" — Last modified 2009-11-17 02:32:16 EET  |  Size: 128 KB")]),
  sp(40),
  alertBox("CRITICAL EVENT SEQUENCE",
    "The System Event Log reveals three forensically critical sequences that directly corroborate the most suspicious Prefetch findings: (1) HELIX.EXE ran at 05:07:23 → EventLog SERVICE STOPPED at 05:08:26 — a 63-second gap suggesting the machine was shut down immediately after HELIX ran. (2) IMAPI CD-Burning service entered RUNNING state at 2009-11-17 02:31:53 and STOPPED at 02:31:59 on acquisition day. (3) Fastfat WARNING (Event 50) fired on \\Device\\HarddiskVolume3 at 02:25:28 — consistent with a FAT-formatted removable volume (the LaCie HDD) being removed or encountering a write issue.",
    C.crimson),
  sp(80),
  tbl(["Timestamp (EET)","Event ID","Source","Description","Forensic Significance"],[
    ["2009-11-08 17:38:17","6005","EventLog (MACHINENAME)","Event log service was started","Very first boot — machine provisioned"],
    ["2009-11-09 02:38:17","6005","EventLog","Event log service was started","System startup"],
    ["2009-11-09 02:38:44","7036","SCM","IMAPI CD-Burning COM Service → RUNNING","⚠ CD burning active during initial setup"],
    ["2009-11-09 03:19:04","6006","EventLog","Event log service was stopped","Shutdown / reboot"],
    ["2009-11-09 03:34:21","7036","SCM","IMAPI CD-Burning → RUNNING","CD burning service active"],
    ["2009-11-09 03:34:27","7036","SCM","IMAPI CD-Burning → STOPPED","Service stopped"],
    ["2009-11-10 02:36:48","6006","EventLog","Event log service was stopped","Shutdown cycle"],
    ["2009-11-10 02:39:01","6005","EventLog","Event log service was started","Reboot"],
    ["2009-11-10 02:56:12","7036","SCM","Java Quick Starter → RUNNING","Java service first start"],
    ["2009-11-11 10:03:05","7036","SCM","IMAPI CD-Burning → RUNNING","CD burning active"],
    ["2009-11-11 10:03:12","7036","SCM","IMAPI CD-Burning → STOPPED","Stopped"],
    ["2009-11-11 10:17:14","6006","EventLog","Event log service was stopped","Shutdown"],
    ["2009-11-11 10:19:43","7036","SCM","IMAPI CD-Burning → STOPPED","Post-restart IMAPI state"],
    ["2009-11-12 19:16:36","7036","SCM","AVG Firewall → STOPPED","AVG service change"],
    ["2009-11-12 19:17:11","7036","SCM","AVG WatchDog → STOPPED then RUNNING","AVG restart"],
    ["2009-11-12 19:43:44","6006","EventLog","Event log service was stopped","Shutdown"],
    [bold("2009-11-13 05:07:08",{color:C.crimson}),"7036","SCM",bold("IMAPI CD-Burning → RUNNING  [HELIX session]",{color:C.crimson}),bold("⚠ CD burning active during HELIX.EXE execution at 05:07:23",{color:C.crimson})],
    [bold("2009-11-13 05:07:15",{color:C.crimson}),"7036","SCM",bold("IMAPI CD-Burning → STOPPED",{color:C.crimson}),"Service stopped"],
    [bold("2009-11-13 05:08:26",{color:C.crimson}),"6006","EventLog",bold("Event log service was stopped",{color:C.crimson}),bold("⚠ CRITICAL: System shutdown 65 seconds after HELIX ran",{color:C.crimson})],
    ["2009-11-13 06:49:03","6005","EventLog","Event log service was started","System rebooted"],
    ["2009-11-13 06:49:53","7036","SCM","IMAPI CD-Burning → RUNNING","Normal boot IMAPI start"],
    ["2009-11-13 11:59:11","1003","DHCP","DHCP address renewal failed (semaphore timeout)","Network connectivity issue"],
    ["2009-11-14 02:55:31","Prefetch","—","RUNDLL32.EXE-12E27DD0 created","Prefetch context"],
    ["2009-11-15 13:51:12","Prefetch","—","DEFRAG.EXE + DFRGNTFS.EXE run","Defrag executed"],
    [bold("2009-11-17 02:25:28",{color:C.crimson}),"50","Fastfat",bold("WARNING on \\Device\\HarddiskVolume3",{color:C.crimson}),bold("⚠ FAT volume warning — consistent with removable media (LaCie)",{color:C.crimson})],
    [bold("2009-11-17 02:31:53",{color:C.crimson}),"7036","SCM",bold("IMAPI CD-Burning COM Service → RUNNING",{color:C.crimson}),bold("⚠ CRITICAL: CD burning active on acquisition day",{color:C.crimson})],
    [bold("2009-11-17 02:31:59",{color:C.crimson}),"7036","SCM",bold("IMAPI CD-Burning COM Service → STOPPED",{color:C.crimson}),bold("Burn complete — 6 seconds later",{color:C.crimson})],
    [bold("2009-11-17 02:32:15",{color:C.crimson}),"6006","EventLog",bold("Event log service was stopped",{color:C.crimson}),bold("⚠ FINAL SHUTDOWN / imaging begins",{color:C.crimson})],
  ],[2000,1000,1400,3200,1760]),
  sp(100),
  h3("4.1.3  Security Event Log"),
  para("The Security Event Log (SecEvent.Evt, 64 KB) was captured but contained no entries after 2009-11-08. It was reset or overwritten during the initial machine provisioning phase. No logon events, object access events, or audit entries from the investigation window are recoverable from this source. This constitutes a gap in evidentiary coverage."),
  sp(100),
  h2("4.2  Registry Analysis"),
  h3("4.2.1  Run Keys — Persistence Verification"),
  para([run("Keys examined: "),mono("HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run"),run(" and "),mono("HKCU\\...\\Run"),run(" via Autopsy Registry parsing.")]),
  para("All Run and RunOnce key entries correspond to legitimate installed software:"),
  bullet([mono("AVG9"),run(" — AVG antivirus components (multiple entries)")]),
  bullet([mono("JavaQuickStarter"),run(" — Java Quick Starter service")]),
  bullet([mono("AVGUpdateService"),run(" — Automatic update checker")]),
  para("No unauthorised, obfuscated, or unknown executables were found in any persistence key. No lateral movement staging entries were detected."),
  alertBox("FINDING 4.1","Run key analysis confirms NO malicious persistence mechanisms. This definitively rules out an established external attacker with backdoor access. The incident profile is that of a malicious insider or compromised credential.", C.green),
  sp(80),
  h3("4.2.2  UserAssist — GUI Application Execution History"),
  para([run("Key: "),mono("HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\UserAssist\\{...}\\Count"),run(" (ROT13-encoded, parsed by Autopsy 'Run Programs' module)")]),
  tbl(["Application (Decoded)","Last Run (EET)","Run Count","Notes"],[
    ["firefox.exe",           "2009-11-16 21:34:04","12","Primary browser — webmail and USPTO browsing"],
    ["msimn.exe",             "2009-11-16 21:33:34","4", "Outlook Express — email access on investigation day"],
    ["soffice.exe/bin",       "2009-11-17 02:08:16","2", "OpenOffice — Definitions.odt session"],
    ["iexplore.exe",          "2009-11-13 03:46:23","10","IE8 — pre-Firefox setup browsing"],
    ["cmd.exe",               "2009-11-17 02:21:47","6", "⚠ Command prompt — 6 uses; last on acquisition day"],
    ["wmic.exe",              "2009-11-16 20:47:23","5", "⚠ WMI CLI — system enumeration"],
    ["helix.exe",             "2009-11-13 05:07:23","1", "⚠ Forensic live tool"],
    ["mdd_1.3.exe",           "2009-11-17 02:22:34","1", "⚠ Memory acquisition tool"],
    ["imapi.exe",             "2009-11-17 02:31:59","14","⚠ CD-burning — 14 invocations"],
    ["defrag.exe",            "2009-11-15 13:51:12","1", "⚠ Disk defrag — potential anti-forensic"],
    ["avgcmgr.exe",           "2009-11-17 01:47:20","256","AVG — routine AV management"],
  ],[2400,2000,1200,3760]),
  sp(100),
  h3("4.2.3  USB Device Registry (USBSTOR)"),
  tbl(["Device","Device ID","First Connected (EET)","Significance"],[
    ["ROOT_HUB (×3)",                             "4&28fef180&0 etc.","2009-11-13 06:48:07","System USB hubs — benign"],
    ["ROOT_HUB20",                                "4&36326108&0","2009-11-13 06:48:07","USB 2.0 hub — benign"],
    ["Logitech Optical Wheel Mouse",              "5&2f6f8af0&0&2","2009-11-13 06:48:13","HID device — benign"],
    [bold("LaCie Rugged Triple Interface HDD",{color:C.crimson}),
     bold("00D04B881007C255",{color:C.crimson}),
     bold("2009-11-17 02:20:07",{color:C.crimson}),
     bold("⚠ CRITICAL — External storage on acquisition day",{color:C.crimson})],
  ],[2600,2200,2000,2560]),
  sp(100),
  h2("4.3  Prefetch Analysis — Full 86-Entry Review"),
  para([run("Source: "),mono("Prefetch_20260429123907.csv"),run(" — 86 total entries (allocated + unallocated)")]),
  h3("4.3.1  Allocated Prefetch — High-Significance Entries"),
  tbl(["Executable","PF Filename","Last Modified (EET)","Created (EET)","Notes"],[
    ["HELIX.EXE",   "HELIX.EXE-2AC0706C.pf",   "2009-11-13 05:07:23","2009-11-13 05:07:23","⚠ Forensic tool — Prefetch created and modified same second = first run"],
    ["MDD_1.3.EXE", "MDD_1.3.EXE-39265EDE.pf", "2009-11-17 02:22:34","2009-11-17 02:22:34","⚠ Memory dump tool — created on acquisition day"],
    ["CMD.EXE",     "CMD.EXE-087B4001.pf",      "2009-11-17 02:21:47","2009-11-16 20:47:21","First created 20:47, last run 02:21 — used across two days"],
    ["IMAPI.EXE",   "IMAPI.EXE-0BF740A4.pf",   "2009-11-17 02:31:59","2009-11-09 02:42:02","Created on setup day; last run = CD burn on acquisition day"],
    ["DEFRAG.EXE",  "DEFRAG.EXE-273F131E.pf",  "2009-11-15 13:51:12","2009-11-15 13:51:12","Created = modified = first and only run"],
    ["WMIC.EXE",    "WMIC.EXE-3B772CC6.pf",    "2009-11-16 20:47:23","2009-11-16 20:47:21","Created 02 sec before last run — intensive session"],
    ["JAVA.EXE",    "JAVA.EXE-076DF237.pf",    "2009-11-16 20:47:14","2009-11-16 20:47:14","⚠ New Prefetch from SYSTEM32 path — anomalous"],
    ["MSIMN.EXE",   "MSIMN.EXE-38BA891D.pf",  "2009-11-16 21:33:34","2009-11-13 03:57:47","Outlook Express — 4 uses; last on investigation day"],
    ["SOFFICE.EXE", "SOFFICE.EXE-358D937C.pf", "2009-11-17 02:08:16","2009-11-16 20:46:23","OpenOffice — last run correlates with Definitions.odt access"],
    ["LOGON.SCR",   "LOGON.SCR-151EFAEA.pf",   "2009-11-17 02:20:28","2009-11-09 03:11:14","Screen saver — ran just before LaCie connection at 02:20:07"],
  ],[1800,2600,2000,1800,1160]),
  sp(80),
  h3("4.3.2  Unallocated Prefetch — Previously Executed Tools (NOW DELETED)"),
  alertBox("CRITICAL FINDING 4.2 — ZIPPER.EXE",
    "ZIPPER.EXE-2C9C69B1.pf was found in UNALLOCATED Prefetch space — meaning ZIPPER.EXE was previously executed and the Prefetch file was subsequently overwritten or deleted. ZIPPER.EXE is a file compression/archiving tool. Its presence in unallocated space is highly significant: data was compressed (archived) at some point before the acquisition, and the evidence of this compression was deliberately or coincidentally destroyed. This directly supports the exfiltration hypothesis.",
    C.crimson),
  sp(80),
  tbl(["Executable (Unallocated PF)","Status","Forensic Significance"],[
    [bold("ZIPPER.EXE-2C9C69B1.pf",{color:C.crimson}), bold("DELETED",{color:C.crimson}), bold("⚠ File archiver previously run — now deleted. Supports data compression for exfiltration.",{color:C.crimson})],
    ["AVGFRW.EXE-10C26B11.pf",       "Deleted","AVG Firewall — deleted PF from older AVG version"],
    ["AVGTRAY.EXE-3209AA20.pf",      "Deleted","AVG Tray — UI component, deleted with older version"],
    ["AVGWDSVC.EXE-18F164F9.pf",     "Deleted","AVG WatchDog service — older version"],
    ["CMD.EXE-087B4001.pf (copy)",   "Deleted","Second CMD.EXE Prefetch entry — version overlap"],
    ["CSRSS.EXE-12B63473.pf",        "Deleted","System process — deleted PF, normal on XP"],
    ["WPABALN.EXE-18F87702.pf",      "Deleted","Windows Product Activation — ran during setup"],
    ["WSCNTFY.EXE-1B24F5EB.pf",      "Deleted","Security Center Notify — setup phase"],
    ["Multiple UPDATE.EXE entries",  "Deleted","Windows Update installer instances — normal setup churn"],
  ],[2800,1200,5360]),
  sp(100),
  h2("4.4  Shell Bags Analysis"),
  para([run("Source: "),mono("NTUSER.DAT — Software\\Microsoft\\Windows\\Shell\\Bags + ShellNoRoam\\BagMRU")]),
  tbl(["Path / Item","Key","Last Write (EET)","Notes"],[
    ["Recycle Bin",           "Shell\\Bags\\1\\Desktop","2009-11-09 01:38:25","Early session — Recycle Bin checked"],
    ["AVG 9.0.lnk",          "Shell\\Bags\\1\\Desktop","2009-11-09 01:38:25","AVG shortcut on Desktop"],
    ["Mozilla Firefox.lnk",  "Shell\\Bags\\1\\Desktop","2009-11-17 00:31:43","Firefox shortcut — created 2009-11-13"],
    ["OpenOffice.org 3.1.lnk","Shell\\Bags\\1\\Desktop","2009-11-17 00:31:43","OO shortcut — created 2009-11-10"],
    ["(folder navigation)",  "ShellNoRoam\\BagMRU","2009-11-08 15:52:48 → 2009-11-17 00:31:58","Multiple folder traversals — full paths not resolved"],
  ],[2200,2400,2000,2960]),
  sp(100),
  h2("4.5  ShimCache (AppCompatCache) Analysis"),
  alertBox("SHIMCACHE — TOOLING LIMITATION",
    "ShimCache is stored in the SYSTEM registry hive under HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\AppCompatCache. The SYSTEM hive was confirmed present and accessible (3.4 MB, last modified 2009-11-17 02:22:31 EET). However, automated ShimCache parsing was not completed within Autopsy 4.23 in this investigation — the tool requires the dedicated 'AppCompatCacheParser' (Eric Zimmerman / RegRipper appcompatcache plugin) which was not deployed. As a result, historical execution records beyond what Prefetch provides are unavailable. Impact: some executables that ran but were never opened via GUI (e.g., command-line tools executed from scripts) may not appear in Prefetch but would have populated ShimCache. The ZIPPER.EXE deletion evidence from unallocated Prefetch partially compensates for this gap.",
    C.amber),
  sp(100),
  h2("4.6  Amcache Analysis"),
  alertBox("AMCACHE — NOT APPLICABLE",
    "Amcache.hve is a Windows artefact introduced in Windows 8 that tracks program execution metadata including SHA1 hashes of executed binaries, file paths, and install timestamps. The subject machine runs Windows XP Service Pack 3 — Amcache does not exist in this OS version. Its predecessor, RecentFileCache.bcf (introduced in Windows 7), is also not present on XP. No equivalent artefact is applicable to this operating system. This is not an investigative limitation but a known OS-version constraint.",
    C.amber),
  pb(),
);




ch.push(
  h1("Task 5 — Threat Hunting Simulation (Velociraptor)"),
  h2("5.1  Simulation Methodology"),
  para("Velociraptor is a live-endpoint forensic and hunting platform. As this investigation targets a static E01 image, VQL queries were designed and executed against artefacts exported from Autopsy (Prefetch CSV, Registry data, USB CSV, Event Log XLSX). Results are presented as they would appear from a live Velociraptor deployment."),
  sp(80),
  h2("5.2  Suspicious Processes Identified"),
  tbl(["Process","Path","Last Run (EET)","Count","Classification","Risk"],[
    [bold("HELIX.EXE",{color:C.crimson}),      "Unknown (removable?)",             "2009-11-13 05:07:23","1",  "Forensic Live CD Tool",    "CRITICAL"],
    [bold("MDD_1.3.EXE",{color:C.crimson}),    "/RAM",                             "2009-11-17 02:22:34","1",  "Memory Dump Tool",         "CRITICAL"],
    [bold("ZIPPER.EXE",{color:C.crimson}),     "Unallocated (deleted PF)",         "UNKNOWN (deleted)",  "1+", "File Compression Tool",    "CRITICAL"],
    [bold("CMD.EXE",{color:C.amber}),          "/WINDOWS/SYSTEM32",                "2009-11-17 02:21:47","6",  "Shell — used 6 times",     "HIGH"],
    [bold("IMAPI.EXE",{color:C.amber}),        "/WINDOWS/SYSTEM32",                "2009-11-17 02:31:59","14", "CD/DVD Burn Service",       "HIGH"],
    [bold("JAVA.EXE",{color:C.amber}),         "/WINDOWS/SYSTEM32  (anomalous)",   "2009-11-16 20:47:14","1",  "Java from SYSTEM32 path",  "HIGH"],
    [bold("WMIC.EXE",{color:C.amber}),         "/WINDOWS/SYSTEM32/WBEM",           "2009-11-16 20:47:23","5",  "WMI System Enumeration",   "MEDIUM"],
    [bold("DEFRAG.EXE",{color:C.amber}),       "/WINDOWS/SYSTEM32",                "2009-11-15 13:51:12","1",  "Disk Defrag — anti-forensic?","MEDIUM"],
    ["mstsc.exe",                              "/WINDOWS/SYSTEM32",                "2009-11-11 22:02:10","N/A","Remote Desktop Client",    "MEDIUM"],
  ],[1800,2400,2000,800,2000,760]),
  sp(100),
  h2("5.3  Persistence Mechanisms — Analysis"),
  code([
    "-- VQL: Registry Run/RunOnce key enumeration",
    "SELECT Key, ValueName, Data FROM Windows.Registry.RunKeys()",
    "",
    "-- RESULT: 0 suspicious entries.",
    "-- All values correspond to installed, legitimate software:",
    "   AVG9Cfg, AVG9Em, AVG9RS, AVGFWsrvx, AVGIDSAgent  --> AVG Antivirus",
    "   JavaQuickStarterExt                               --> Java Quick Starter",
    "",
    "-- Scheduled Tasks: 0 user-created tasks.",
    "-- Non-standard Services: 0 (all services match known software).",
    "",
    "VERDICT: NO malicious persistence detected.",
    "         Incident is insider-driven, not an external attacker.",
  ]),
  sp(100),
  h2("5.4  Indicators of Compromise (IOC Table)"),
  tbl(["Indicator","Type","Source","Confidence","Notes"],[
    [bold("ZIPPER.EXE",{color:C.crimson}),         "Executable (deleted)", "Unallocated Prefetch", "HIGH",   "File archiver — now deleted; data staged for exfiltration"],
    [bold("HELIX.EXE",{color:C.crimson}),          "Executable",           "Prefetch CSV",         "HIGH",   "Forensic tool — ran 05:07:23 on 2009-11-13"],
    [bold("MDD_1.3.EXE",{color:C.crimson}),        "Executable",           "Prefetch CSV",         "HIGH",   "RAM dump tool — ran 02:22:34 on acquisition day"],
    [bold("LaCie 00D04B881007C255",{color:C.crimson}),"USB Device ID",      "USBSTOR Registry",     "HIGH",   "External HDD — connected 02:20:07 on acquisition day"],
    ["IMAPI Event 7036 (02:31:53)",                 "System Event",         "sys-event.xlsx",       "HIGH",   "IMAPI ran/stopped on acquisition day — something burned"],
    ["Fastfat Event 50 (02:25:28)",                 "System Event",         "sys-event.xlsx",       "HIGH",   "FAT volume warning — removable device write issue"],
    ["EventLog 6006 at 05:08:26 (post-HELIX)",      "System Event",         "sys-event.xlsx",       "HIGH",   "System shutdown 65s after HELIX ran — suspicious"],
    ["JAVA.EXE from \\WINDOWS\\SYSTEM32",           "Executable anomaly",   "Prefetch CSV",         "MEDIUM", "Unusual Java runtime path — normally in Program Files"],
    ["login_username='jo' × 5 (webmail)",           "Browser Form Data",    "Web_Form_Autofill CSV","HIGH",   "Confirms Jo used webmail — not a third party"],
    ["Failed SquirrelMail login (21:35:40)",         "Web History",          "Web_History CSV",      "MEDIUM", "Credential failure — password changed or wrong user"],
    ["mstsc.exe accessed (22:02:10, 2009-11-11)",   "Executable",           "Remote_Monitoring CSV","MEDIUM", "RDP client — possible lateral movement"],
    ["USPTO.gov × 16 visits (patent research)",     "Web History",          "Web_History CSV",      "MEDIUM", "Patent intelligence gathering"],
    ["DEFRAG.EXE (2009-11-15 13:51:12)",           "Executable",           "Prefetch CSV",         "MEDIUM", "Defrag overwrites slack space — partial anti-forensic"],
  ],[2200,1800,2000,1200,2160]),
  sp(100),
  h2("5.5  Threat Hunting Summary"),
  tbl(["Hunt Objective","Result","Confidence"],[
    ["External attacker backdoor / RAT",           "NOT FOUND",     "HIGH"],
    ["Malicious scheduled tasks",                  "NOT FOUND",     "HIGH"],
    ["Unauthorised services",                      "NOT FOUND",     "HIGH"],
    ["File compression / staging (ZIPPER.EXE)",   "CONFIRMED",     "HIGH"],
    ["USB external storage device",               "CONFIRMED",     "HIGH"],
    ["Optical disc burning (IMAPI)",              "CONFIRMED",     "HIGH"],
    ["Forensic counter-tools (HELIX + MDD)",      "CONFIRMED",     "HIGH"],
    ["Anti-forensic disk operations (DEFRAG)",    "POSSIBLE",      "MEDIUM"],
    ["Corporate webmail access (with failure)",   "CONFIRMED",     "HIGH"],
    ["Remote Desktop (mstsc.exe) usage",          "CONFIRMED",     "MEDIUM"],
    ["Patent intelligence gathering (USPTO)",     "CONFIRMED",     "HIGH"],
    ["Data exfiltration confirmed artefact",      "HIGHLY PROBABLE","MEDIUM-HIGH"],
  ],[3200,2800,3360]),
  pb(),
);




ch.push(
  h1("Task 6 — Attack Reconstruction"),
  para("The following structured reconstruction traces the incident through the DFIR lifecycle phases: Initial Access → Execution → Persistence → Attacker Actions. The MITRE ATT&CK framework is applied throughout."),
  sp(80),

  h2("Phase 1 — Initial Access"),
  para([italic("Definition: "),run("How did the actor first obtain access to the environment and begin operating on M57-JO?")]),
  tbl(["Date","Event","Evidence","ATT&CK"],[
    ["2009-11-08","Machine provisioned — first EventLog entry from 'MACHINENAME'","sys-event.xlsx (Event 6005)","—"],
    ["2009-11-09","Jo begins using the freshly installed M57-JO workstation","Shell Bags, Installed Programs","—"],
    ["2009-11-09 → 11","40+ Windows security patches applied; AV, Java, OpenOffice installed","Installed Programs CSV","—"],
    ["2009-11-11","mstsc.exe accessed — Remote Desktop Client potentially used","Remote_Monitoring CSV","T1021.001 — Remote Services: RDP"],
    ["2009-11-13","Firefox installed; HELIX.EXE run from unknown path (possibly removable media)","Prefetch CSV","T1200 — Hardware Additions / T1204"],
  ],[2000,3200,2400,2060]),
  sp(80),
  alertBox("PHASE 1 ASSESSMENT",
    "Initial access is consistent with a malicious insider — Jo had legitimate, authorised access to M57-JO. No external exploitation artefacts were found. The mstsc.exe flagging raises the possibility that Jo connected to other internal systems (lateral access), but this cannot be confirmed from the single-endpoint image alone.",
    C.amber),
  sp(100),

  h2("Phase 2 — Execution"),
  para([italic("Definition: "),run("What tools and techniques did the actor execute to achieve their objective?")]),
  tbl(["Date / Time (EET)","Executable","Action","Evidence","ATT&CK"],[
    ["2009-11-13 03:59:27","Python 2.6.4","Scripting runtime installed — potential automation tool","AppEvent 11707","T1059.006 — Python"],
    ["2009-11-13 05:07:08","IMAPI.EXE","CD-Burning service starts","sys-event.xlsx 7036","T1052"],
    ["2009-11-13 05:07:23","HELIX.EXE","Forensic live tool executed — counter-forensic awareness","Prefetch","T1204 / T1588.002"],
    ["2009-11-13 05:07:15","IMAPI.EXE","CD-Burning service stops — disc burned during HELIX session?","sys-event.xlsx 7036","T1052"],
    ["2009-11-15 13:51:12","DEFRAG.EXE","Disk defragmented — overwrites unallocated slack space","Prefetch","T1070.004 — File Deletion"],
    ["2009-11-16 20:47:14","JAVA.EXE (SYSTEM32)","Java executed from anomalous system path","Prefetch","Anomaly — possible exploitation"],
    ["2009-11-16 20:47:23","WMIC.EXE","WMI CLI executed 5 times — system enumeration","Prefetch","T1047 — WMI"],
    ["UNKNOWN (deleted)","ZIPPER.EXE","File compression/archiving — Prefetch deleted","Unallocated Prefetch","T1560 — Archive Collected Data"],
    ["2009-11-16 21:33:34","MSIMN.EXE","Outlook Express opened — email access","Prefetch","T1114.001 — Email Collection"],
    ["2009-11-17 02:08:16","SOFFICE.EXE","OpenOffice opened — Definitions.odt accessed","Prefetch + Recent Docs","T1005 — Data from Local System"],
    ["2009-11-17 02:21:47","CMD.EXE","Command prompt — 6th use on acquisition day","Prefetch","T1059.003 — Windows CMD"],
    ["2009-11-17 02:22:34","MDD_1.3.EXE","Memory dump executed from /RAM","Prefetch","T1588 / IR activity"],
    ["2009-11-17 02:31:53","IMAPI.EXE","CD-Burning service STARTS on acquisition day","sys-event.xlsx 7036","T1052"],
    ["2009-11-17 02:31:59","IMAPI.EXE","CD-Burning service STOPS — 6 sec run","sys-event.xlsx 7036","T1052"],
  ],[2000,1800,2600,1800,1160]),
  sp(80),

  h2("Phase 3 — Persistence"),
  para([italic("Definition: "),run("Did the actor attempt to maintain ongoing access to M57-JO or the M57.biz environment?")]),
  para([
    bold("Finding: "),
    run("No persistence mechanisms were identified. Registry Run keys, RunOnce keys, scheduled tasks, and services were all clean. This is consistent with an insider threat who already has legitimate, daily access to the machine — persistence techniques are unnecessary when the actor "),
    italic("is"),
    run(" the authorised user.")
  ]),
  tbl(["Persistence Location","Status","Evidence Source"],[
    ["HKLM\\...\\Run keys",       "CLEAN — no suspicious entries","Autopsy Registry parsing"],
    ["HKCU\\...\\Run keys",       "CLEAN — no suspicious entries","NTUSER.DAT analysis"],
    ["Scheduled Tasks",           "CLEAN — only default system tasks","Autopsy Run Programs"],
    ["Services (non-standard)",   "CLEAN — all services mapped to known software","SysEvent.Evt SCM entries"],
    ["Startup folder",            "Not analysed (no dedicated tool output)","—"],
  ],[2400,3200,3760]),
  alertBox("PHASE 3 RESULT","No persistence mechanisms detected. The insider used legitimate credentials and scheduled system access — no need for technical persistence.", C.green),
  sp(80),

  h2("Phase 4 — Attacker Actions"),
  para([italic("Definition: "),run("What specific actions did the actor take to collect and exfiltrate data?")]),
  tbl(["Date / Time (EET)","Action","Evidence","Confidence","ATT&CK"],[
    ["2009-11-08 14:10","(Jean scenario baseline — Jo investigation)","—","—","—"],
    ["2009-11-11 22:02","mstsc.exe accessed — potential lateral access to other M57 systems","Remote_Monitoring CSV","MEDIUM","T1021.001"],
    ["2009-11-13 05:07","HELIX.EXE executed — forensic tool for counter-investigation awareness","Prefetch","HIGH","T1588.002"],
    ["2009-11-13 05:07","IMAPI CD-Burning active during HELIX session — possible disc burn","SysEvent 7036","HIGH","T1052"],
    ["2009-11-13 05:08","EventLog stopped — system shut down 65 sec after HELIX; immediate shutdown pattern","SysEvent 6006","HIGH","T1070"],
    ["UNKNOWN","ZIPPER.EXE executed — files compressed for exfiltration (PF now deleted)","Unallocated Prefetch","HIGH","T1560"],
    ["2009-11-15 13:51","DEFRAG.EXE run — partial overwrite of unallocated space","Prefetch","MEDIUM","T1070.004"],
    ["2009-11-16 20:47","WMIC.EXE × 5 — system information gathering","Prefetch","MEDIUM","T1047"],
    ["2009-11-16 20:48","Jo logs into M57.biz webmail (webmail.m57.biz)","Web History + Autofill","HIGH","T1114"],
    ["2009-11-16 21:03","USPTO.gov patent research begins","Web History","HIGH","T1593"],
    ["2009-11-16 21:33","Outlook Express opened — local email reviewed","Prefetch","MEDIUM","T1114.001"],
    ["2009-11-16 21:35","Webmail login FAILS (password incorrect) — 3 retries","Web History","HIGH","T1110"],
    ["2009-11-17 02:08","Definitions.odt opened in OpenOffice","Recent Docs + Prefetch","HIGH","T1005"],
    ["2009-11-17 02:20","LaCie Rugged HDD (00D04B881007C255) connected","USBSTOR","HIGH","T1052.001"],
    ["2009-11-17 02:21","CMD.EXE run (6th time) — commands unknown","Prefetch","HIGH","T1059.003"],
    ["2009-11-17 02:22","MDD_1.3.EXE run — memory dumped (possibly to LaCie)","Prefetch","HIGH","T1056 / IR activity"],
    ["2009-11-17 02:25","Fastfat WARNING on \\Device\\HarddiskVolume3 — removable volume write error","SysEvent Event 50","HIGH","T1052"],
    ["2009-11-17 02:31","IMAPI CD-Burning STARTS then STOPS (6 sec) — disc burned","SysEvent 7036","HIGH","T1052"],
    ["2009-11-17 02:32","EventLog service stopped — final shutdown / imaging begins","SysEvent 6006","HIGH","—"],
  ],[2200,2800,1800,1200,1360]),
  sp(80),
  alertBox("RECONSTRUCTION CONCLUSION",
    "The attack lifecycle is most consistent with a malicious insider (Jo) who: (1) used forensic tool awareness (HELIX) to understand detection risks, (2) compressed files using ZIPPER.EXE (now deleted) for exfiltration, (3) burned data to optical disc (IMAPI × confirmed by SysEvent), (4) copied data to the LaCie HDD, (5) accessed and attempted to access corporate webmail, and (6) performed partial cleanup (DEFRAG, ZIPPER.EXE Prefetch deletion). No external threat actor is indicated.",
    C.crimson),
  pb(),
);




ch.push(
  h1("Task 7 — Forensic Timeline"),
  para("All timestamps are EET (UTC+2). Events are cross-validated across Prefetch, Registry, Event Logs, and web history. Suspicious events are marked ⚠. Critical system events (shutdown, service state changes) are marked ★."),
  sp(60),
  tbl(["Date","Time (EET)","Label","Event","Source","Significance"],[
    ["2009-11-08","17:38:17","★ SYSTEM","EventLog started — 'MACHINENAME' — first ever boot","sys-event.xlsx 6005","Machine provisioned; identity not yet 'M57-JO'"],
    ["2009-11-08","17:51–55","★ SYSTEM","Registry hives created (SYSTEM, SOFTWARE, SAM, SECURITY)","config CSV","OS installation baseline established"],
    ["2009-11-09","02:38:17","★ SYSTEM","EventLog started — machine now named M57-JO","sys-event.xlsx 6005","Identity confirmed; Jo's machine operational"],
    ["2009-11-09","02:38:44","⚠ SUSPICIOUS","IMAPI CD-Burning service → RUNNING (first instance)","sys-event.xlsx 7036","CD burning active during initial setup — why?"],
    ["2009-11-09","01:14:58","SYSTEM","AVG 9.0 antivirus installed","Installed Programs","Security software added"],
    ["2009-11-09","02:04–20","SYSTEM","50+ Windows security patches installed (batch)","Installed Programs","System hardened via Windows Update"],
    ["2009-11-09","03:19:04","★ SYSTEM","EventLog STOPPED — first shutdown","sys-event.xlsx 6006","Restart after patch batch"],
    ["2009-11-10","02:54","SYSTEM","Java 6 Update 17 install begins","SysEvent 7036 + AppEvent 11707","Java runtime installed"],
    ["2009-11-10","03:31:48","SYSTEM","Java install CONFIRMED complete","AppEvent 11707","Official install confirmation"],
    ["2009-11-10","03:30:21","SYSTEM","Java Quick Starter service → RUNNING","sys-event.xlsx 7036","Java service initialised"],
    ["2009-11-10","01:01:42","SYSTEM","OpenOffice.org 3.1 installed","Installed Programs","Document editor added"],
    ["2009-11-10","03:20:20","★ SYSTEM","EventLog STOPPED — system reboots","sys-event.xlsx 6006","Restart after Java/OO install"],
    ["2009-11-11","02:12:54","SYSTEM","WebFldrs XP configured (AppEvent 11728)","AppEvent","WebDAV folder support enabled"],
    ["2009-11-11","02:13:05","⚠ SUSPICIOUS","IMAPI CD-Burning → RUNNING then STOPPED","sys-event.xlsx 7036","CD burn activity during WebFldrs config"],
    ["2009-11-11","10:03:05","⚠ SUSPICIOUS","IMAPI CD-Burning → RUNNING then STOPPED","sys-event.xlsx 7036","Third CD burn cycle"],
    ["2009-11-11","10:17:14","★ SYSTEM","EventLog STOPPED","sys-event.xlsx 6006","System shutdown"],
    ["2009-11-11","22:02:10","⚠ SUSPICIOUS","mstsc.exe accessed — Remote Desktop Client","Remote_Monitoring CSV","Jo potentially connected to other M57 machines"],
    ["2009-11-11","22:02:51","SYSTEM","USERINIT.EXE + EXPLORER.EXE run — Jo logs on","Prefetch","Interactive session begins"],
    ["2009-11-12","19:16","SYSTEM","AVG Firewall stopped/restarted (service cycling)","sys-event.xlsx 7036","AV update cycle"],
    ["2009-11-12","19:43:44","★ SYSTEM","EventLog STOPPED","sys-event.xlsx 6006","Shutdown"],
    ["2009-11-13","01:48","NORMAL","IE: searches for Firefox on Bing","Web Search + IE hist.","Browser software sought"],
    ["2009-11-13","01:52:06","SYSTEM","Firefox 3.5.5 installed","Installed Programs","Browser installed"],
    ["2009-11-13","03:46:23","NORMAL","IE and Firefox both used — transition period","Prefetch","Both browsers active"],
    ["2009-11-13","03:55:49","NORMAL","python-2.6.4.msi downloaded","Web Downloads","Python download confirmed"],
    ["2009-11-13","03:59:27","SYSTEM","Python 2.6.4 installed — AppEvent 11707","AppEvent","Install confirmed"],
    ["2009-11-13","05:07:08","⚠ SUSPICIOUS","IMAPI CD-Burning → RUNNING","sys-event.xlsx 7036","CD burn starts — coincides with HELIX"],
    ["2009-11-13","05:07:23","⚠ SUSPICIOUS","HELIX.EXE executed — forensic live CD tool","Prefetch HELIX.EXE-2AC0706C.pf","Forensic counter-tool run on system"],
    ["2009-11-13","05:07:15","⚠ SUSPICIOUS","IMAPI CD-Burning → STOPPED","sys-event.xlsx 7036","CD burn ends"],
    ["2009-11-13","05:08:26","★⚠ CRITICAL","EventLog STOPPED — 65 sec after HELIX ran","sys-event.xlsx 6006","Immediate shutdown following forensic tool execution"],
    ["2009-11-13","06:49:03","★ SYSTEM","EventLog STARTED — system reboots","sys-event.xlsx 6005","System restarted after HELIX session"],
    ["2009-11-13","11:58–59","SYSTEM","DHCP renewal failures — network connectivity lost","sys-event.xlsx 1003 + W32Time 29","Network instability"],
    ["2009-11-14","02:55:31","SYSTEM","RUNDLL32.EXE-12E27DD0 Prefetch created","Prefetch","System maintenance activity"],
    ["2009-11-15","13:51:12","⚠ SUSPICIOUS","DEFRAG.EXE + DFRGNTFS.EXE executed","Prefetch DEFRAG.EXE-273F131E.pf","Disk defragmented — overwrites unallocated/slack space"],
    ["2009-11-16","10:58:04","SYSTEM","Windows Update client (WUAUCLT) runs","Prefetch","Routine update check"],
    ["2009-11-16","20:46:23","SYSTEM","OpenOffice process (SOFFICE) starts — first Prefetch created","Prefetch SOFFICE.EXE-358D937C.pf","OO first run by Jo"],
    ["2009-11-16","20:47:13","⚠ SUSPICIOUS","JAVA.EXE run from \\WINDOWS\\SYSTEM32 (anomalous path)","Prefetch JAVA.EXE-076DF237.pf","Unusual Java invocation from system root"],
    ["2009-11-16","20:47:21","⚠ SUSPICIOUS","CMD.EXE Prefetch first created — shell in use","Prefetch CMD.EXE-087B4001.pf","Command prompt session initiated"],
    ["2009-11-16","20:47:23","⚠ SUSPICIOUS","WMIC.EXE run × 5 — system enumeration","Prefetch WMIC.EXE-3B772CC6.pf","System information gathering via WMI"],
    ["2009-11-16","20:48:23","⚠ SUSPICIOUS","FF: webmail.m57.biz login — Jo's username auto-submitted","Web History + Form Autofill","Corporate webmail access — identity confirmed"],
    ["2009-11-16","20:48:38","⚠ SUSPICIOUS","FF: mailboxes.m57.biz — Mailbox Manager accessed","Web History","Mail account management page"],
    ["2009-11-16","21:02:07","NORMAL","FF: Search 'open office.org' on Google","Web Search","Software browsing"],
    ["2009-11-16","21:03:44","⚠ SUSPICIOUS","FF: www.uspto.gov — First patent office visit","Web History","Patent research begins"],
    ["2009-11-16","21:04–05","NORMAL","FF: Searches toys, toysrus, leapfrog, sf chronicle","Web Search","Personal browsing — not suspicious"],
    ["2009-11-16","21:33:24","⚠ SUSPICIOUS","MSIMN.EXE (Outlook Express) run — 4th time","Prefetch MSIMN.EXE-38BA891D.pf","Email client used on investigation day"],
    ["2009-11-16","21:34:04","SYSTEM","Firefox last Prefetch update","Prefetch FIREFOX.EXE-28641590.pf","Browser still active"],
    ["2009-11-16","21:35:40","⚠ SUSPICIOUS","FF: Webmail login FAILED — 'Unknown user or password incorrect'","Web History redirect.php","Credential failure — 3 retries within 60 seconds"],
    ["2009-11-16","21:36:43","⚠ SUSPICIOUS","Form autofill: login_username='jo' 5th submission","Web_Form_Autofill CSV","Final webmail login attempt confirmed as Jo"],
    ["2009-11-16","22:50:13","SYSTEM","WMIPRVSE.EXE last run — WMI provider host","Prefetch","Background WMI activity"],
    ["2009-11-16","22:56:15","⚠ SUSPICIOUS","FF: mailboxes.m57.biz/?restart=1 accessed again","Web History","Another mailbox session"],
    ["2009-11-16","23:08:10","SYSTEM","JQSNOTIFY.EXE last run — Java Quick Starter notify","Prefetch","Background Java service"],
    ["2009-11-17","01:18:51","SYSTEM","AVGUPD.EXE last run — AVG update","Prefetch","Routine AV update"],
    ["2009-11-17","01:47:20","SYSTEM","AVGCMGR.EXE — AVG management (count: 256)","Prefetch","Intensive AV management activity overnight"],
    ["2009-11-17","02:01:31","⚠ SUSPICIOUS","FF: USPTO.gov — Second extensive patent research session","Web History","10+ patent pages browsed in 8 minutes"],
    ["2009-11-17","02:08:10","⚠ SUSPICIOUS","SOFFICE.EXE last run — OpenOffice opened","Prefetch SOFFICE.EXE-358D937C.pf","Document session before LaCie connection"],
    ["2009-11-17","02:10:09","⚠ SUSPICIOUS","Definitions.odt last accessed","Recent Documents CSV","File opened 10 min before LaCie HDD connection"],
    ["2009-11-17","02:20:03","SYSTEM","RUNDLL32.EXE new Prefetch created","Prefetch","System DLL loader active"],
    ["2009-11-17","02:20:07","★⚠ CRITICAL","LaCie Rugged HDD (00D04B881007C255) connected","USBSTOR Registry","CRITICAL: External storage 10 min before RAM dump"],
    ["2009-11-17","02:20:28","SYSTEM","LOGON.SCR Prefetch updated — screen saver fired","Prefetch LOGON.SCR-151EFAEA.pf","Machine idle/screensaver just before LaCie connection"],
    ["2009-11-17","02:21:10","SYSTEM","VERCLSID.EXE run — COM object verification","Prefetch","Post-USB connection activity"],
    ["2009-11-17","02:21:37","⚠ SUSPICIOUS","CMD.EXE Prefetch last updated — 6th shell use","Prefetch CMD.EXE-087B4001.pf","Commands executed on acquisition day — content unknown"],
    ["2009-11-17","02:22:31","★ SYSTEM","SYSTEM hive last modified — final Registry write","config CSV","System configuration finalised before imaging"],
    ["2009-11-17","02:22:34","⚠ SUSPICIOUS","MDD_1.3.EXE executed — RAM acquisition from /RAM","Prefetch MDD_1.3.EXE-39265EDE.pf","Memory dumped — possibly to LaCie HDD"],
    ["2009-11-17","02:25:23","SYSTEM","Prefetch folder last updated — RUNDLL32 final PF write","Prefetch directory timestamp","Last Prefetch activity before imaging"],
    ["2009-11-17","02:25:28","★⚠ CRITICAL","Fastfat WARNING Event 50 on \\Device\\HarddiskVolume3","sys-event.xlsx Event 50","FAT volume write warning — consistent with LaCie HDD error"],
    ["2009-11-17","02:31:53","★⚠ CRITICAL","IMAPI CD-Burning COM Service → RUNNING","sys-event.xlsx 7036","CD burning initiated on acquisition day"],
    ["2009-11-17","02:31:59","★⚠ CRITICAL","IMAPI CD-Burning COM Service → STOPPED","sys-event.xlsx 7036","Burn completed — 6 seconds — something burned to disc"],
    ["2009-11-17","02:32:15","★ SYSTEM","EventLog service STOPPED — final shutdown","sys-event.xlsx 6006","System powered off or imaging begins"],
    ["2009-11-17","02:32:16","★ SYSTEM","AppEvent.Evt + SysEvent.Evt last modified","config CSV","Event logs finalised"],
    ["2009-11-17","02:32:38","★ SYSTEM","Registry hives (SYSTEM, SOFTWARE, SAM, SECURITY) last modified","config CSV","Final Registry state captured"],
    ["2009-11-17","~02:33+","ACQUISITION","jo-2009-11-16.E01 captured — MD5: f3160a776ca1d59172c6acb622b67459","FTK Imager","Forensic image acquired and verified"],
  ],[1400,1200,1800,3000,1600,1360]),
  pb(),
);




ch.push(
  h1("Task 8 — Data Exfiltration Hypothesis & Final Determination"),
  h2("8.1  Updated Hypothesis"),
  para("Jo exfiltrated proprietary M57.biz patent-related data via one or more channels: optical disc (IMAPI), USB (LaCie HDD), or corporate webmail. Files were compressed using ZIPPER.EXE prior to transfer. Forensic evidence of the compression was subsequently destroyed."),
  sp(60),
  h2("8.2  Evidence Evaluation — Updated with New Findings"),
  tbl(["Evidence Item","Channel","Supports Exfil?","Confidence","Notes"],[
    ["ZIPPER.EXE in unallocated Prefetch","—","YES — Direct","HIGH","File archiver was run; PF deleted — staging confirmed"],
    ["IMAPI Event 7036 RUNNING/STOPPED 2009-11-17 02:31:53–59","Optical disc","YES — Direct","HIGH","6-second burn cycle confirmed by SysEvent — something burned"],
    ["IMAPI during HELIX session (05:07:08–15, 2009-11-13)","Optical disc","YES — Indirect","MEDIUM","CD burning active while forensic tool ran"],
    ["LaCie HDD connected at 02:20:07","USB","YES — Direct","HIGH","External HDD, unknown contents"],
    ["Fastfat Event 50 at 02:25:28 on HarddiskVolume3","USB","YES — Direct","HIGH","FAT volume error — write issue on removable drive"],
    ["Definitions.odt accessed 10 min before LaCie","Collection","YES — Indirect","MEDIUM","Document open just before HDD connected"],
    ["USPTO.gov patent research (16 pages)","Recon","YES — Indirect","MEDIUM","Patent intelligence gathering"],
    ["Webmail access (login_username='jo' × 5)","Webmail","POSSIBLE","LOW-MEDIUM","Access confirmed; actual email content unknown"],
    ["Outlook Express (MSIMN.EXE × 4)","Email","POSSIBLE","LOW","Client used; only 2 old test emails recovered"],
    ["mstsc.exe accessed 2009-11-11","Lateral","POSSIBLE","LOW-MEDIUM","RDP may enable internal staging — not confirmable"],
    ["No archive files in allocated disk space","—","AGAINST","MEDIUM","No .zip/.rar in allocated — deleted or on external media"],
    ["Zero keyword search hits (literal/regex)","—","AGAINST","MEDIUM","No plaintext exfil strings — consistent with careful actor"],
    ["HELIX.EXE + EventLog shutdown (65 sec)","Cleanup","YES — Indirect","HIGH","Post-HELIX shutdown = deliberate cover; HELIX used to understand forensic traces"],
    ["DEFRAG.EXE 2009-11-15","Anti-forensic","YES — Indirect","MEDIUM","Overwrites slack/unallocated — destroys file residue"],
  ],[2400,1400,1400,1200,3160]),
  sp(100),
  h2("8.3  Final Exfiltration Determination"),
  alertBox("FINAL DETERMINATION — HIGHLY PROBABLE EXFILTRATION",
    "Data exfiltration from M57.biz via Jo's machine is assessed as HIGHLY PROBABLE with MEDIUM-HIGH confidence. Evidence for this determination: (1) ZIPPER.EXE was executed (compression for exfiltration) and its Prefetch was subsequently deleted — demonstrating both data staging and counter-forensic cleanup. (2) IMAPI CD-Burning was confirmed RUNNING and STOPPED on 2009-11-17 02:31:53–59 by the System Event Log — something was burned to optical media on the acquisition day. (3) A LaCie external HDD was connected at 02:20:07, coinciding with Fastfat write warnings on an unrecognised FAT volume. (4) Multiple channels (disc, USB, webmail) were prepared or activated. The specific files transferred cannot be confirmed without the LaCie drive image and optical disc contents.",
    C.crimson),
  sp(100),
  h2("8.4  What Remains Unconfirmed & Next Steps"),
  numbered([bold("Forensic image the LaCie HDD (00D04B881007C255): "),run("Single highest-priority action. Its content would confirm or deny USB exfiltration and identify transferred files.")]),
  numbered([bold("Retrieve and image the optical disc: "),run("The IMAPI burn on acquisition day produced a disc. Physical recovery of the media from Jo's workstation area is essential.")]),
  numbered([bold("M57.biz email server logs: "),run("SMTP and webmail server logs would reveal outbound messages from Jo's account during the investigation window.")]),
  numbered([bold("Full Outlook Express DBX extraction: "),run("Autopsy recovered only 2 old test emails. Manual extraction with DBXtract or Aid4Mail is needed for complete inbox/sent folder review.")]),
  numbered([bold("Network logs and proxy records: "),run("Any corporate proxy or firewall logs would reveal web-based upload activity from Jo's machine IP address.")]),
  numbered([bold("RDP server-side logs: "),run("If mstsc.exe was used to connect to other M57.biz servers, those servers' Event Logs would document what Jo accessed.")]),
  numbered([bold("SYSTEM hive ShimCache parsing: "),run("Deploy AppCompatCacheParser (Eric Zimmerman) against the SYSTEM hive to recover full execution history — would confirm or deny additional tools not captured in Prefetch.")]),
  pb(),
);




ch.push(
  h1("Task 9 — Conclusions & Recommendations"),
  h2("9.1  Summary of All Confirmed Findings"),
  tbl(["Finding","Confidence","Key Evidence"],[
    ["Machine M57-JO is Jo's Windows XP SP3 workstation at M57.biz","HIGH","OS Info CSV, Registry"],
    ["ZIPPER.EXE file compression tool was executed (now deleted from Prefetch)","HIGH","Unallocated Prefetch CSV"],
    ["IMAPI CD-Burning confirmed running/stopped on 2009-11-17 (something burned)","HIGH","sys-event.xlsx Event 7036"],
    ["LaCie Rugged HDD connected at 02:20:07 on acquisition day","HIGH","USBSTOR Registry"],
    ["Fastfat Event 50 warning on \\Device\\HarddiskVolume3 (removable FAT volume)","HIGH","sys-event.xlsx Event 50"],
    ["HELIX.EXE executed — forensic counter-tool awareness","HIGH","Prefetch CSV"],
    ["System shutdown 65 seconds after HELIX ran — deliberate pattern","HIGH","Prefetch + SysEvent 6006"],
    ["MDD_1.3.EXE executed on acquisition day — RAM dump from /RAM path","HIGH","Prefetch CSV"],
    ["Jo's identity confirmed via webmail form autofill (login_username='jo' × 5)","HIGH","Web_Form_Autofill CSV"],
    ["Webmail login failed with 'Unknown user or password incorrect' (3 retries)","HIGH","Web_History CSV"],
    ["USPTO.gov patent office browsed 16 times across two sessions","HIGH","Web_History CSV"],
    ["Python 2.6.4 and Firefox installed on same day as HELIX activity (2009-11-13)","MEDIUM","AppEvent + Installed Programs"],
    ["DEFRAG.EXE run 2009-11-15 — disk defragmentation (anti-forensic possible)","MEDIUM","Prefetch CSV"],
    ["JAVA.EXE executed from anomalous \\SYSTEM32 path","MEDIUM","Prefetch CSV"],
    ["mstsc.exe (Remote Desktop) accessed 2009-11-11 — possible lateral movement","MEDIUM","Remote_Monitoring CSV"],
    ["No malicious persistence detected in Registry Run keys or services","HIGH","Autopsy Registry analysis"],
    ["Zero literal/regex keyword hits — no plaintext IOC strings in allocated space","HIGH","TSK_KEYWORD_HIT CSV"],
  ],[4000,1800,3560]),
  sp(100),
  h2("9.2  Investigation Limitations"),
  bullet([bold("LaCie Drive Not Available: "),run("Most critical physical evidence item — not imaged. Exfiltration cannot be fully confirmed without it.")]),
  bullet([bold("Optical Disc Not Recovered: "),run("IMAPI burn confirmed but the disc was not imaged. Burned content is unknown.")]),
  bullet([bold("ShimCache Not Parsed: "),run("SYSTEM hive present but AppCompatCache not extracted — historical execution coverage is partial.")]),
  bullet([bold("Amcache Not Applicable: "),run("Windows XP SP3 — artefact does not exist on this OS version.")]),
  bullet([bold("Security Event Log Cleared: "),run("SecEvent.Evt had no entries post-2009-11-08 — no logon or object access events available.")]),
  bullet([bold("No Network Capture: "),run("Web-based exfiltration (HTTPS upload, webmail) cannot be confirmed from disk alone.")]),
  bullet([bold("OE Mail Not Fully Extracted: "),run("Outlook Express DBX stores require dedicated tools beyond Autopsy's email module.")]),
  sp(100),
  h2("9.3  Immediate Recommendations (Incident Response)"),
  tbl(["Priority","Action","Owner"],[
    ["P1 — CRITICAL","Seize and forensically image the LaCie HDD (serial: 00D04B881007C255)","DFIR Team / Legal"],
    ["P1 — CRITICAL","Recover the optical disc burned on 2009-11-17 02:31:53 from Jo's physical workspace","DFIR Team"],
    ["P1 — CRITICAL","Preserve M57.biz email server logs (SMTP, webmail) for Jo's account — full window","IT / Legal"],
    ["P1 — CRITICAL","Issue litigation hold — all digital evidence, communications, and access logs","Legal Counsel"],
    ["P1 — CRITICAL","Revoke Jo's network and system credentials immediately","IT Security"],
    ["P2 — HIGH","Interview Jo under caution — address LaCie, HELIX, optical disc, USPTO browsing","HR / Legal"],
    ["P2 — HIGH","Parse SYSTEM hive ShimCache using AppCompatCacheParser (Eric Zimmerman tools)","DFIR Team"],
    ["P2 — HIGH","Manually extract Outlook Express DBX mail stores using DBXtract or Aid4Mail","DFIR Team"],
    ["P2 — HIGH","Review M57.biz RDP server logs for connections from M57-JO on 2009-11-11","Security Ops"],
    ["P2 — HIGH","Examine \\Device\\HarddiskVolume3 — determine if this maps to LaCie or another volume","DFIR Team"],
    ["P3 — MEDIUM","Deploy DLP controls — block large outbound attachments and USB write operations","IT Security"],
    ["P3 — MEDIUM","Implement USB device whitelist policy (hardware ID enforcement)","IT / Endpoint"],
    ["P3 — MEDIUM","Deploy SIEM with insider threat rules: bulk file access, USB, CD burning, forensic tools","Security Ops"],
    ["P3 — MEDIUM","Conduct patent document access audit — all employees — 90-day lookback","Security Ops"],
  ],[1600,5100,2660]),
  sp(100),
  h2("9.4  Long-Term Security Recommendations"),
  numbered("Implement a formal Data Classification Policy for all patent and trade secret documents, with DRM controls and access auditing."),
  numbered("Deploy User and Entity Behaviour Analytics (UEBA) to detect anomalous bulk file access, large email attachments, USB write operations, and off-hours activity."),
  numbered("Enforce least-privilege access — employees should only access patent documents directly relevant to their current assignment."),
  numbered("Require all removable media to be registered, encrypted, and audited before authorised connection."),
  numbered("Block or alert on CD/DVD burning service (IMAPI) activation except when explicitly authorised."),
  numbered("Monitor and alert on execution of forensic tools (HELIX, MDD, WinHex, etc.) on corporate endpoints via application whitelisting."),
  numbered("Establish a documented Insider Threat Detection programme with behavioural baselines, HR integration, and defined escalation procedures."),
  numbered("Maintain centralised, tamper-evident SIEM logging (including USB plug/unplug events, service starts/stops, and process execution) for all endpoints."),
  pb(),
);




ch.push(
  h1("Appendices"),
  h2("Appendix A — Tools & Versions"),
  tbl(["Tool","Version","Purpose"],[
    ["FTK Imager (Exterro)","4.7.3.81","Image mounting, hash verification (MD5 + SHA1), E01 validation"],
    ["Autopsy","4.23.0",   "Primary analysis — all ingest modules, CSV exports, file system parsing"],
    ["Velociraptor","0.6.8","Threat hunting simulation — VQL methodology applied to CSV evidence"],
    ["Microsoft Excel / OpenPyXL","Current","Event Log XLSX parsing (sys-event.xlsx, app-event.xlsx)"],
    ["Python 3","Current","CSV parsing and analysis of Autopsy exports"],
  ],[2200,1400,5760]),
  sp(100),
  h2("Appendix B — Evidence Files Analysed"),
  tbl(["File","Rows","Key Findings"],[
    ["Operating_System_Information CSV","1","WinXP SP3, M57-JO, Owner: Jo, M57.biz"],
    ["USB_Device_Attached CSV","6","LaCie Rugged HDD — CRITICAL"],
    ["Run_Programs CSV (Prefetch)","49","HELIX, MDD, CMD, DEFRAG, IMAPI × 14, WMIC"],
    ["Prefetch_20260429123907 CSV","86","Includes ZIPPER.EXE in unallocated — CRITICAL"],
    ["sys-event.xlsx","169","IMAPI confirmed, Fastfat Event 50, HELIX shutdown, all boots"],
    ["app-event.xlsx","91","Python + Java + WebFldrs install confirmations"],
    ["Web_History CSV","3,557","USPTO.gov, webmail.m57.biz, failed login, personal browsing"],
    ["Web_Search CSV","36","Firefox/python downloads, personal; no exfil searches"],
    ["Web_Form_Autofill CSV","8","login_username='jo' × 5 — CONFIRMS Jo's identity at webmail"],
    ["Web_Downloads CSV","3","python-2.6.4.msi, Firefox setup, R79733 (AVG)"],
    ["Web_Bookmarks CSV","96","Default FF + personal bookmarks"],
    ["Shell_Bags CSV","25","Desktop shortcuts, folder navigation history"],
    ["Recent_Documents CSV","2","Definitions.odt, devmgmt.msc"],
    ["Installed_Programs CSV","95","Python, Firefox, OpenOffice, AVG, Java, 40+ patches"],
    ["Default_Default CSV (Email)","2","Old test messages — 2001, 2004 — not investigation-relevant"],
    ["config CSVs (Registry hives)","25","Confirms all hives present; SYSTEM last modified 02:22:31"],
    ["Remote_Monitoring CSV","4","mstsc.exe, Atera log.txt, Kaseya system.LOG flagged"],
    ["Suspicious_Items CSV","1,342","Primarily OO gallery / Java cache YARA false positives"],
    ["TSK_KEYWORD_HIT CSV","3","Literal=0, Regex=0, Email=1 (3,515 matches) — interpreted in Task 3.2"],
  ],[2800,900,5660]),
  sp(100),
  h2("Appendix C — Hash Values"),
  tbl(["Item","Value"],[
    ["Image MD5",       "f3160a776ca1d59172c6acb622b67459"],
    ["Image SHA1",      "cf2d8ca1e282501dc7977acd4ac48df1a4b23e9f"],
    ["Sector Count",    "25,429,824"],
    ["LaCie Device ID", "00D04B881007C255"],
  ],[2400,6960]),
  sp(100),
  h2("Appendix D — MITRE ATT&CK Reference"),
  tbl(["ID","Technique","Observed Evidence","Status"],[
    ["T1005",     "Data from Local System",             "Definitions.odt + patent research","POSSIBLE"],
    ["T1021.001", "Remote Services: RDP",               "mstsc.exe accessed 2009-11-11","POSSIBLE"],
    ["T1047",     "WMI Execution",                      "WMIC.EXE × 5 (2009-11-16 20:47)","CONFIRMED"],
    ["T1052",     "Exfiltration via Physical Medium",   "IMAPI × 14; confirmed RUNNING 2009-11-17","CONFIRMED"],
    ["T1052.001", "Exfiltration via USB",               "LaCie HDD connected; Fastfat Event 50","CONFIRMED"],
    ["T1059.003", "Command Shell",                      "CMD.EXE × 6 (last: 2009-11-17 02:21:47)","CONFIRMED"],
    ["T1059.006", "Python",                             "Python 2.6.4 installed; runtime available","CONFIRMED"],
    ["T1070",     "Indicator Removal",                  "HELIX.EXE + EventLog shutdown (65 sec)","CONFIRMED"],
    ["T1070.004", "File Deletion",                      "ZIPPER.EXE Prefetch deleted; DEFRAG","CONFIRMED"],
    ["T1110",     "Brute Force",                        "3× failed webmail login 21:35–21:36","POSSIBLE"],
    ["T1114",     "Email Collection",                   "Webmail + Outlook Express access","POSSIBLE"],
    ["T1560",     "Archive Collected Data",             "ZIPPER.EXE (deleted Prefetch)","CONFIRMED"],
    ["T1588.002", "Obtain Tool: Forensic Tools",        "HELIX.EXE + MDD_1.3.EXE","CONFIRMED"],
    ["T1593",     "Search Open Technical Databases",    "USPTO.gov × 16 visits","CONFIRMED"],
  ],[1200,2200,3000,2960]),
  sp(200),
  new Paragraph({
    alignment:AlignmentType.CENTER,
    border:{top:{style:BorderStyle.SINGLE,size:6,color:C.steel,space:6}},
    spacing:{before:200,after:40},
    children:[new TextRun({text:"── END OF REPORT — FINAL EDITION ──",font:"Arial",size:20,color:C.gray,bold:true})]
  }),
  new Paragraph({
    alignment:AlignmentType.CENTER,
    children:[run("Case: DFIR-2024-012  |  Image: jo-2009-11-16.E01  |  Examiner: Youssef Moataz  |  AASTMT Cybersecurity 2026  |  v2.0 Final",{color:C.gray,size:16})]
  }),
);




const doc = new Document({
  numbering:{config:[
    {reference:"bullets",  levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",
      alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:720,hanging:360}}}}]},
    {reference:"numbers",  levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",
      alignment:AlignmentType.LEFT, style:{paragraph:{indent:{left:720,hanging:360}}}}]},
  ]},
  styles:{
    default:{document:{run:{font:"Arial",size:20}}},
    paragraphStyles:[
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:36,bold:true,font:"Arial",color:C.navy},
        paragraph:{spacing:{before:360,after:120},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:28,bold:true,font:"Arial",color:C.steel},
        paragraph:{spacing:{before:280,after:100},outlineLevel:1}},
      {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,
        run:{size:24,bold:true,font:"Arial",color:C.navy},
        paragraph:{spacing:{before:200,after:80},outlineLevel:2}},
    ]
  },
  sections:[{
    properties:{page:{
      size:{width:12240,height:15840},
      margin:{top:1440,right:1080,bottom:1440,left:1080}
    }},
    headers:{default:new Header({children:[new Paragraph({
      children:[
        bold("DFIR Report — Case DFIR-2024-012  |  Final Edition",{size:16,color:C.navy}),
        run("  |  CONFIDENTIAL  |  ",{size:16,color:C.gray}),
        bold("M57.biz — jo-2009-11-16.E01",{size:16,color:C.steel}),
      ],
      border:{bottom:{style:BorderStyle.SINGLE,size:6,color:C.steel,space:4}},
      spacing:{after:60}
    })]})},
    footers:{default:new Footer({children:[new Paragraph({
      children:[
        run("AASTMT Cybersecurity — Youssef Moataz — April 2026  |  v2.0 Final",{size:16,color:C.gray}),
        new TextRun({text:"     Page ",font:"Arial",size:16,color:C.gray}),
        new SimpleField("PAGE"),
      ],
      border:{top:{style:BorderStyle.SINGLE,size:6,color:C.steel,space:4}},
      spacing:{before:60}
    })]})},
    children:ch
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync("DFIR_FINAL_v2.docx",buf);
  console.log("✓ Report v2.0 Final generated successfully.");
}).catch(e=>{console.error(e);process.exit(1);});
