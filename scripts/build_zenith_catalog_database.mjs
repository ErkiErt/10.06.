import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.join("outputs", "zenith_catalog_database_20260609");
const outputPath = path.join(outputDir, "Zenith_Kataloogi_Tootebaas_2026-06-09.xlsx");
const sourceExtractPath = path.join(outputDir, "source_extract.json");
const catalogSourceFile = "Zenith catalogue 2020_0.pdf";
const catalogSourceId = "ZENITH_CAT_2020";
const businessRule =
  "Tööloogika: Plastoki oma materjalibaas on esmane kontroll. Zenithist tellitakse juurde siis, kui Plastokil pole sobivat materjali, mõõtu, sertifikaati, kogust või oma laoseis/materjal ei kata kliendi vajadust.";
let sourceExtract = {
  catalog_file: catalogSourceFile,
  catalog_page_count: 0,
  catalog_pages: [],
  price_rows: [],
  price_summary: [],
  discounts: [],
  plastok_rows: [],
  plastok_summary: [],
  plastok_top_categories: [],
  plastok_top_groups: [],
  plastok_top_materials: [],
};

try {
  sourceExtract = JSON.parse(await fs.readFile(sourceExtractPath, "utf8"));
} catch {
  // The workbook can still be built from the catalog-only data if source extracts
  // have not been generated yet.
}

const zenithPriceRows = sourceExtract.price_rows ?? [];
const zenithPriceSummaryRows = sourceExtract.price_summary ?? [];
const zenithDiscountRows = sourceExtract.discounts ?? [];
const plastokRows = sourceExtract.plastok_rows ?? [];
const plastokSummaryRows = sourceExtract.plastok_summary ?? [];
const catalogPages = new Map((sourceExtract.catalog_pages ?? []).map((row) => [row.page, row.text ?? ""]));

function sourceText(page) {
  return `${catalogSourceFile}, lk ${page}`;
}

function normArticle(value) {
  return String(value ?? "").toUpperCase().replace(/\s+/g, " ").trim();
}

function compactArticle(value) {
  return normArticle(value).replace(/\s+/g, "");
}

function auditCompact(value) {
  return String(value ?? "")
    .toUpperCase()
    .replace(/[^A-Z0-9ÕÄÖÜŠŽ]/g, "");
}

function pageText(page) {
  return catalogPages.get(page) ?? "";
}

function pageContains(page, value) {
  const needle = auditCompact(value);
  if (!needle) return false;
  return auditCompact(pageText(page)).includes(needle);
}

function ocrVariants(value) {
  const base = auditCompact(value);
  const variants = new Set();
  const replacements = [
    ["S", "5"],
    ["I", "L"],
    ["L", "I"],
    ["O", "0"],
    ["0", "O"],
    ["Y", "V"],
  ];
  for (const [from, to] of replacements) {
    variants.add(base.replaceAll(from, to));
    for (let index = 0; index < base.length; index += 1) {
      if (base[index] === from) {
        variants.add(`${base.slice(0, index)}${to}${base.slice(index + 1)}`);
      }
    }
  }
  return [...variants].filter((variant) => variant && variant !== base);
}

function pageContainsOcrVariant(page, value) {
  if (pageContains(page, value)) return false;
  const text = auditCompact(pageText(page));
  return ocrVariants(value).some((variant) => text.includes(variant));
}

function fmtNum(value, decimals = 2) {
  if (typeof value !== "number" || Number.isNaN(value)) return "";
  return Number(value.toFixed(decimals));
}

function getAny(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null) return row[key];
  }
  return "";
}

function includesAny(value, terms) {
  const text = String(value ?? "").toUpperCase();
  return terms.some((term) => text.includes(term));
}

const plastokFamilyProfiles = {
  "EPDM kumm": {
    role: "Oma ilmastiku-, osooni- ja veekindel kummileht",
    uses: "Välitingimused; tihendid; vesi; UV/osoon; üldine kummileht",
    strengths: "Ilmastiku- ja osoonikindel; hea välitingimustesse; lai temperatuurivahemik",
    cautions: "Ei ole esimene valik õli-, rasva- ega kütusekeskkonda.",
    zenithTrigger: "Kui vaja konkreetset Zenith mõõtu, joogivee erilahendust, Manticore seeria sertifikaati või Plastoki oma mõõt/laoseis ei sobi.",
  },
  "NBR kumm": {
    role: "Oma õli- ja rasvakindlam kummileht",
    uses: "Õli ja rasv; töökojad; tihendid; kütusekeskkond",
    strengths: "Õli- ja rasvakindlam kui EPDM/SBR; sobib paljudesse tööstus- ja töökojakeskkondadesse",
    cautions: "UV, osoon ja püsiv välitingimus vajavad eraldi kontrolli.",
    zenithTrigger: "Kui vaja Zenith Satyr/Pegasus NBR kindlat mõõtu, sertifikaati või hinnakirja artiklit.",
  },
  "SBR/NR üld- ja kulumiskumm": {
    role: "Oma üld-, kulumis- ja sahakummi baas",
    uses: "Üldine tööstus; kulumiskaitse; konveier; sahakumm; tihendid",
    strengths: "Hea kulumis- ja löögitaluvus sõltuvalt kvaliteedist; lai mõõduvalik",
    cautions: "Õli/rasv ja tugev UV/ozone keskkond vajavad NBR/CR/EPDM kontrolli.",
    zenithTrigger: "Kui oma kulumiskummi mõõt, kõvadus või kvaliteet ei kata vajadust, kontrolli Zenith Mining & Para / Red Dragoon / Typhon.",
  },
  "CR / neopreen": {
    role: "Oma ilmastiku ja mõõduka õlikindlusega erikumm",
    uses: "Välitingimused; offshore; mõõdukas õli/kemikaal; tööstus",
    strengths: "Kompromiss EPDM-i ja NBR-i vahel: ilmastik + mõõdukas õlikindlus",
    cautions: "Toidukontakt ja erikemikaalid vajavad sertifikaati või FKM/PTFE kontrolli.",
    zenithTrigger: "Kui vaja Wendigo/CR konkreetset Zenith mõõtu või Plastoki CR valik on liiga kitsas.",
  },
  "Silikoon": {
    role: "Oma kõrge/madala temperatuuri ja toidukontakti materjal",
    uses: "Toiduainetööstus; kõrge temperatuur; madal temperatuur; tihendid",
    strengths: "Väga lai temperatuurivahemik; paljud FDA/toidukontakti variandid",
    cautions: "Ei ole kulumiskumm; mehaaniline rebimine ja abrasioon vajavad kontrolli.",
    zenithTrigger: "Kui vaja Viking seeria värvi, kõvadust, laiust või Zenith hinnakirja artiklit.",
  },
  "FKM / Viton": {
    role: "Oma kõrge temperatuuri ja keemiakindluse erimaterjal",
    uses: "Kõrge temperatuur; kemikaalid; õlid; süsivesinikud; eritingimused",
    strengths: "Kõrge kemikaali-, õli- ja temperatuuritaluvus",
    cautions: "Kallis erimaterjal; konkreetne kemikaal tuleb alati üle kontrollida.",
    zenithTrigger: "Kui Plastokil sobiv FKM mõõt puudub, kontrolli Zenith Karibu FKM.",
  },
  "Butüül / IR": {
    role: "Oma gaasitiheduse ja vibratsioonisummutuse materjal",
    uses: "Tihendid; gaasitihedus; vibratsioon; löögisummutus",
    strengths: "Väga madal gaasiläbilaskvus; hea vibratsiooni ja löögi summutamiseks",
    cautions: "Õli- ja kütusekeskkond vajab eraldi kontrolli.",
    zenithTrigger: "Kui oma butüülmõõt puudub või vaja Rama Butyl hinnakirja artiklit.",
  },
  "Kummimatid": {
    role: "Oma põranda-, libisemisvastaste ja kaitsemattide baas",
    uses: "Põrand; libisemisvastane pind; tööalad; haagised; sissepääsud",
    strengths: "Praktiline laomaterjal põrandale, kaitseks ja libisemise vähendamiseks",
    cautions: "Elektriisolatsioon ja toidukontakt vajavad eraldi sertifikaati.",
    zenithTrigger: "Kui mustrit, laiust, paksust või erimatti Plastokil pole, kontrolli Zenith matting seeriat.",
  },
  "Elektriisolatsiooni matid": {
    role: "Oma dielektriliste mattide kontroll",
    uses: "Elektrikilbid; elektriisolatsioon; ohutuspõrandad",
    strengths: "Sobib elektriohu vähendamise põrandakatteks, kui klass/sertifikaat sobib",
    cautions: "Pingeklass ja sertifikaat tuleb alati tellimuse/projekti järgi kinnitada.",
    zenithTrigger: "Kui oma dielektriline matt ei kata pingeklassi/mõõtu, kontrolli Feniks Insulation Mat.",
  },
  "Vibratsioonimaterjalid": {
    role: "Oma vibratsiooni- ja mürasummutuse materjalid",
    uses: "Masinaalused; vibratsioon; mürasummutus",
    strengths: "Vähendab vibratsiooni ja müra; sobib masinaalusteks sõltuvalt koormusest",
    cautions: "Koormus kg/cm2, paksus ja deformatsioon tuleb üle kontrollida.",
    zenithTrigger: "Kui oma VIBRA/VIBRAFOAM ei sobi, kontrolli Zenith Novibramat või Rama Butyl.",
  },
  "PTFE / tihendiplaadid": {
    role: "Oma keemia- ja tihendimaterjalide baas",
    uses: "Tihendid; kemikaalid; madal hõõrdumine; äärikutihendid",
    strengths: "Hea keemiakindlus ja tihendirakendused sõltuvalt materjalist",
    cautions: "Mehaaniline koormus, temperatuur ja tihendistandard tuleb kinnitada.",
    zenithTrigger: "Zenith kumm ei asenda tihendiplaati; kasuta Zenithit ainult siis, kui vaja kummilehte.",
  },
  "PE / UHMW plastid": {
    role: "Oma plastmaterjal kulumisele, juhikutele ja lõikelaudadele",
    uses: "Juhikud; kulumisdetailid; lõikelauad; toiduainetööstuse plastid",
    strengths: "Hea kulumis- ja libisemisomadus; PE1000/UHMW sobib paljudeks juhikuteks",
    cautions: "Ei ole kumm; elastsus, tihendamine ja temperatuur erinevad kummidest.",
    zenithTrigger: "Zenith kumm ainult siis, kui klient vajab elastset kummilahendust, mitte plastikut.",
  },
  "PA / nailon": {
    role: "Oma tugev tehniline plast",
    uses: "Puksid; rullid; detailid; mehaanilised osad",
    strengths: "Tugev, töödeldav ja kulumiskindel tehniline plast",
    cautions: "Niiskusimavus ja mõõdupüsivus vajavad kontrolli.",
    zenithTrigger: "Zenith kumm ei asenda PA detaili; võrdle ainult funktsiooni, mitte materjali järgi.",
  },
  "POM / atsetaal": {
    role: "Oma mõõdupüsiv tehniline plast",
    uses: "Täpsemad detailid; puksid; liugdetailid; mehaanika",
    strengths: "Mõõdupüsiv, töödeldav ja hea liugomadusega",
    cautions: "Ei ole elastne tihendi- või kummimaterjal.",
    zenithTrigger: "Zenith kumm ainult siis, kui kasutus nõuab elastsust või kummipinda.",
  },
  "PUR / PU": {
    role: "Oma väga kulumis- ja rebimiskindel polüuretaan",
    uses: "Kulumisosad; rattad; kaabitsad; löögikaitse",
    strengths: "Väga hea kulumis- ja rebimiskindlus",
    cautions: "Temperatuur, kemikaalid ja kõvadus vajavad kontrolli.",
    zenithTrigger: "Kui vaja kummilehte või suuremat painduvat rulli, kontrolli Zenith Mining & Para.",
  },
  "Muu plast": {
    role: "Muu Plastoki plastmaterjal",
    uses: "Detailid; lehed; vardad; torud",
    strengths: "Oma plastibaas katab palju tehnilisi detaile",
    cautions: "Sobivus sõltub konkreetsest plastist.",
    zenithTrigger: "Zenith ainult juhul, kui lahendus peab olema kummist.",
  },
  "Muu kumm": {
    role: "Muu Plastoki kummimaterjal",
    uses: "Tihendid; lehed; profiilid; üldkasutus",
    strengths: "Võib katta kliendi vajaduse ilma Zenithist tellimata",
    cautions: "Materjalikood ja omadused tuleb konkreetse rea järgi üle kontrollida.",
    zenithTrigger: "Kui oma kummireal puudub sobiv materjal/mõõt/sertifikaat, kontrolli Zenithit.",
  },
  "Määramata / kontrollida": {
    role: "Materjal vajab käsitsi kontrolli",
    uses: "Sisemised artiklid, eridetailid või puuduliku materjalikoodiga read",
    strengths: "Võib olla vajalik olemasolev lao- või kliendispetsiifiline artikkel",
    cautions: "Ära kasuta automaatseks materjalisoovituseks ilma rea kontrollita.",
    zenithTrigger: "Kui materjali ei saa tuvastada, küsi täpsustus või kontrolli Zenithit ainult kasutusvajaduse põhjal.",
  },
};

function normalizePlastokFamily(row) {
  const material = `${row["Material"] ?? ""}`;
  const name = `${row["Name"] ?? ""}`;
  const category = `${row["Category"] ?? ""}`;
  const group = `${row["Group"] ?? ""}`;
  const text = `${material} ${name} ${category} ${group}`.toUpperCase();
  if (includesAny(text, ["DIELEKTR", "ELS30000", "ELS50000"])) return "Elektriisolatsiooni matid";
  if (includesAny(text, ["VIBRA", "ANTIVIB", "ANTI-VIBRA"])) return "Vibratsioonimaterjalid";
  if (includesAny(text, ["FKM", "VITON"])) return "FKM / Viton";
  if (includesAny(text, ["SILIKOON", "SILICONE", "VMQ"])) return "Silikoon";
  if (includesAny(text, ["BUTYL", "BUTÜÜL", " IR ", "RAMA BUTYL"])) return "Butüül / IR";
  if (includesAny(text, ["KULUTUS", "SAHK", "NRKUMM", "NR/SBR", "SBR", "NR "])) return "SBR/NR üld- ja kulumiskumm";
  if (includesAny(text, ["EPDM"])) return "EPDM kumm";
  if (includesAny(text, ["NBR"])) return "NBR kumm";
  if (includesAny(text, ["NEOPREEN", " CR", "CR,", "KALIPSO"])) return "CR / neopreen";
  if (includesAny(text, ["MATT", "MATS", "RÕNGAS", "RUUT", "MÜNT", "APELSIN", "CHECKER"])) return "Kummimatid";
  if (includesAny(text, ["PTFE", "TESNIT", "TEMASIL", "TIHEN"])) return "PTFE / tihendiplaadid";
  if (includesAny(text, ["PE1000", "PE500", "PE300", "PE100", "PE,", " PE "])) return "PE / UHMW plastid";
  if (includesAny(text, ["PA6", "PA66", "PA,", " PA "])) return "PA / nailon";
  if (includesAny(text, ["POM"])) return "POM / atsetaal";
  if (includesAny(text, ["PUR", "PU,", " PU ", "POLÜURETAAN"])) return "PUR / PU";
  if (category.includes("Plastic") || includesAny(group, ["TPLAS", "PE", "PP", "PVC", "PMMA", "PC", "PET", "PEEK", "PVDF", "PCTFE", "ABS"])) return "Muu plast";
  if (category.includes("Rubber") || includesAny(group, ["KUMM", "TKUMM"])) return "Muu kumm";
  return "Määramata / kontrollida";
}

const groupDescriptions = [
  ["SBR rubber sheets", "NR/SBR sheets for general industrial use, wear, shock absorption, gaskets, insulation and conveyor belt related applications.", "p5-p9"],
  ["NBR rubber sheets", "Oil and grease resistant sheets with good mechanical properties; used in chemical industries, garages and workshops.", "p11-p16"],
  ["EPDM rubber sheets", "Weather, ozone, oxygen and cold resistant sheets; suitable for outdoor, automotive, hot water, solar and potable water applications depending on grade.", "p17-p24"],
  ["CR rubber sheets", "Chloroprene sheets with moderate oil, ozone and weather resistance; suitable for offshore and industrial use.", "p25-p30"],
  ["Mining & Para", "Wear and impact resistant rubber sheeting, pulley lagging, dust sealing and repair products for mining, conveyor and severe abrasion applications.", "p31-p44"],
  ["Food Grade rubber sheets", "Food-contact rubber sheets compliant with FDA and/or EC1935/2004 where stated in the catalog.", "p45-p52"],
  ["FKM rubber sheets", "High temperature and chemical resistant FKM for demanding industrial applications.", "p53-p55"],
  ["Various", "Specialty rubber sheets including NBR/PVC, CSM and butyl rubber.", "p57-p60"],
  ["Rubber matting", "Ribbed, checker, coin, diamond, insulation and anti-vibration matting products.", "p61-p71"],
];

function codes(prefix, suffixes) {
  return suffixes.map((suffix) => `${prefix} ${suffix}`);
}

function p(row) {
  return {
    productId: row.productId,
    group: row.group,
    product: row.product,
    page: row.page,
    quality: row.quality ?? "",
    color: row.color ?? "",
    tempRange: row.tempRange ?? "",
    specificGravity: row.specificGravity ?? "",
    shoreA: row.shoreA ?? "",
    tensileMpa: row.tensileMpa ?? "",
    elongationPct: row.elongationPct ?? "",
    abrasion: row.abrasion ?? "",
    certification: row.certification ?? "",
    finish: row.finish ?? "",
    thicknesses: row.thicknesses ?? "",
    widths: row.widths ?? "",
    lengths: row.lengths ?? "",
    features: row.features ?? "",
    application: row.application ?? "",
    articleNumbers: row.articleNumbers ?? [],
    note: row.note ?? "",
  };
}

const gp = {
  sbr: "SBR rubber sheets",
  nbr: "NBR rubber sheets",
  epdm: "EPDM rubber sheets",
  cr: "CR rubber sheets",
  mining: "Mining & Para",
  food: "Food Grade rubber sheets",
  fkm: "FKM rubber sheets",
  various: "Various",
  matting: "Rubber matting",
};

const commonGeneral = "Widely used industrial rubber quality with good mechanical properties, abrasion and shock resistance.";
const generalUse = "General purpose gaskets, shock absorbers, conveyor belts, insulation etc.";
const nbrEcoFeature = "Widely used oil and grease resistant industrial rubber quality with good mechanical properties.";
const nbrUse = "Chemical industries, garages, workshops etc.";
const epdmFeature = "Weather and cold resistant rubber with good mechanical properties and chemical resistance; suitable for open air and ozone exposure.";
const epdmUse = "Automotive industry, outdoor use and solar panels.";
const crFeature = "Chloroprene industrial rubber with good mechanical properties, moderate oil and ozone resistance.";
const crUse = "Offshore and industry.";
const paraUse = "Very wide range of applications, conveyor systems, pumps etc.";
const matBasicFeature = "Anti-slip ribbed surface on one side and a fine cloth impression pattern on the other side.";
const matBasicUse = "Loose lay applications for floors, walkways, runners, loading areas, workbenches and general-purpose areas.";

const products = [
  p({ productId: "SX70150", group: gp.sbr, product: "Black Miami", page: 7, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.5 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "3", elongationPct: "200", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12; 15; 20; 25; 30; 40; 50 mm", widths: "1.2/1.4 m; 1.4 m", lengths: "5/10/15/20 m", features: commonGeneral, application: generalUse, articleNumbers: codes("SX70150", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "10/0-1x1", "12/0", "15/0", "15/0-1x1", "20/0", "20/0-1x1", "25/0", "25/0-1x1", "30/0", "30/0-1x1", "40/0-1x1", "50/0-1x1"]) }),
  p({ productId: "SB65130", group: gp.sbr, product: "Zenstar", page: 8, quality: "NR/SBR", color: "Black", tempRange: "-40 to +75 deg C", specificGravity: "1.3 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "6", elongationPct: "350", abrasion: "250 mm3", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m; 2 m on selected thicknesses", lengths: "5/10/15/20 m", features: commonGeneral, application: generalUse, articleNumbers: codes("SB65130", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "SX50150", group: gp.sbr, product: "Black Miami 50", page: 9, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.5 grs/cm3", shoreA: "50 Shore A +/- 5", tensileMpa: "3", elongationPct: "200", thicknesses: "2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10 m", features: commonGeneral, application: generalUse, articleNumbers: codes("SX50150", ["2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),

  p({ productId: "NB7010", group: gp.nbr, product: "Satyr Eco", page: 13, quality: "NBR/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "4", elongationPct: "200", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12; 15; 20; 30; 40; 50 mm", widths: "1.2/1.4 m; 1.4 m", lengths: "2/5/10/15/20 m", features: nbrEcoFeature, application: nbrUse, articleNumbers: codes("NB7010", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "12/0", "15/0", "20/0", "30/0", "40/0", "50/0"]) }),
  p({ productId: "NB70140", group: gp.nbr, product: "Satyr Medium", page: 14, quality: "NBR/SBR", color: "Black", tempRange: "-25 to +80 deg C", specificGravity: "1.45 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "5", elongationPct: "250", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.2/1.4 m; 1.4 m", lengths: "5/10/15/20 m", features: "Oil and grease resistant industrial rubber. Resistant to gasoline, oil, propane, natural gas and coal gas.", application: nbrUse, articleNumbers: codes("NB70140", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "NB65130", group: gp.nbr, product: "Satyr Premium", page: 15, quality: "NBR", color: "Black", tempRange: "-30 to +110 deg C", specificGravity: "1.35 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "10", elongationPct: "350", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/15/20 m", features: "Excellent oil and grease resistant industrial rubber quality. Resistant to oil, super, pentane and bio-diesel.", application: nbrUse, articleNumbers: codes("NB65130", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "NB5010", group: gp.nbr, product: "Satyr 50", page: 16, quality: "NBR/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "50 Shore A +/- 5", tensileMpa: "4", elongationPct: "200", thicknesses: "1; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/20 m", features: nbrEcoFeature, application: nbrUse, articleNumbers: codes("NB5010", ["1/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),

  p({ productId: "EP7010", group: gp.epdm, product: "Manticore Eco", page: 19, quality: "EPDM/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "4", elongationPct: "200", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12; 15; 20; 25; 30; 40; 50 mm", widths: "1.2/1.4 m; 1.4 m", lengths: "2/5/10/15/20 m", features: epdmFeature, application: epdmUse, articleNumbers: codes("EP7010", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "12/0", "15/0", "20/0", "25/0", "30/0", "40/0", "50/0"]) }),
  p({ productId: "EP65130", group: gp.epdm, product: "Manticore Medium", page: 20, quality: "EPDM/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.35 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "6", elongationPct: "250", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/15/20 m", features: epdmFeature, application: epdmUse, articleNumbers: codes("EP65130", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "EP65120", group: gp.epdm, product: "Manticore Premium", page: 21, quality: "EPDM", color: "Black", tempRange: "-40 to +130 deg C", specificGravity: "1.25 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "8", elongationPct: "350", thicknesses: "1; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/20 m", features: "Excellent weather and cold resistant industrial rubber with good mechanical properties and chemical resistance.", application: "Automotive industry, hot water seals, outdoor use and solar panels.", articleNumbers: codes("EP65120", ["1/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "EP5010", group: gp.epdm, product: "Manticore 50", page: 22, quality: "EPDM/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "50 Shore A +/- 5", tensileMpa: "4", elongationPct: "200", thicknesses: "2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10 m", features: epdmFeature, application: epdmUse, articleNumbers: codes("EP5010", ["2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "EP70110W", group: gp.epdm, product: "Manticore Aqua", page: 23, quality: "EPDM", color: "Black", tempRange: "-40 to +120 deg C", specificGravity: "1.15 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "11.5", elongationPct: "350", thicknesses: "2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10 m", features: "Potable water rubber sheet according to different European standards; certificates available on request.", application: "Water fitting parts such as gaskets, membranes and washers used in contact with drinking water.", articleNumbers: codes("EP70110W", ["2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "EP70122PC", group: gp.epdm, product: "Manticore Peroxid", page: 24, quality: "EPDM", color: "Black", tempRange: "-40 to +140 deg C", specificGravity: "1.25 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "9", elongationPct: "200", thicknesses: "1; 2; 3; 4; 5 mm", widths: "1.4 m", lengths: "10/20 m", features: "Peroxide cured EPDM with high temperature range, chemical resistance and good ozone resistance.", application: "Automotive industry, hot water seals, outdoor use and solar panels.", articleNumbers: codes("EP70122PC", ["1/0", "2/0", "3/0", "4/0", "5/0"]) }),

  p({ productId: "CR7010", group: gp.cr, product: "Wendigo Eco", page: 27, quality: "CR/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "4", elongationPct: "200", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12; 15; 20; 25; 30; 40; 50 mm", widths: "1.4 m", lengths: "2/5/10/20 m", features: crFeature, application: crUse, articleNumbers: codes("CR7010", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "10/0-1x1", "12/0", "15/0", "15/0-1x1", "20/0", "20/0-1x1", "25/0", "25/0-1x1", "30/0", "40/0", "50/0"]) }),
  p({ productId: "NE70140", group: gp.cr, product: "Wendigo Medium", page: 28, quality: "CR/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "5", elongationPct: "250", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/15/20 m", features: crFeature, application: crUse, articleNumbers: codes("NE70140", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "NE65140", group: gp.cr, product: "Wendigo Premium", page: 29, quality: "CR", color: "Black", tempRange: "-25 to +100 deg C", specificGravity: "1.4 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "10", elongationPct: "400", thicknesses: "1; 2; 3; 4; 5; 6; 10 mm", widths: "1.4 m", lengths: "5/10/20 m", features: "Excellent chloroprene rubber with resistance to heat, ozone, weathering and oil; flame retardant properties.", application: crUse, articleNumbers: codes("NE65140", ["1/0", "2/0", "3/0", "4/0", "5/0", "6/0", "10/0"]) }),
  p({ productId: "CR5010", group: gp.cr, product: "Wendigo 50", page: 30, quality: "CR/SBR", color: "Black", tempRange: "-30 to +90 deg C", specificGravity: "1.45 grs/cm3", shoreA: "50 Shore A +/- 5", tensileMpa: "4", elongationPct: "200", thicknesses: "2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10 m", features: crFeature, application: crUse, articleNumbers: codes("CR5010", ["2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),

  p({ productId: "AB45110", group: gp.mining, product: "Typhon", page: 33, quality: "NR (Para)", color: "Beige, Grey", tempRange: "-30 to +70 deg C", specificGravity: "1.10 grs/cm3", shoreA: "45 Shore A +/- 5", tensileMpa: "15", elongationPct: "600", abrasion: "100 mm3", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12; 15; 20 mm", widths: "1.4 m", lengths: "5/10/15/20 m", features: "Natural rubber quality with good mechanical properties, abrasion resistance, low temperature flexibility and excellent adhesion.", application: paraUse, articleNumbers: codes("AB45110", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "12/0", "15/0", "20/0"]) }),
  p({ productId: "AB45110R", group: gp.mining, product: "Centaur Red (BL)", page: 34, quality: "NR (Para)", color: "Red", tempRange: "-30 to +70 deg C", specificGravity: "1.10 grs/cm3", shoreA: "45 Shore A +/- 5", tensileMpa: "15", elongationPct: "600", abrasion: "100 mm3", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12; 15; 20 mm", widths: "1.4/2 m; 2 m on selected thicknesses", lengths: "10/15/20 m", features: "Natural rubber quality with abrasion resistance and excellent adhesion; bonding layer available on selected variants.", application: paraUse, articleNumbers: codes("AB45110R", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "12/0", "15/0", "20/0"]) }),
  p({ productId: "ABE45120R", group: gp.mining, product: "Centaur Red Eco", page: 35, quality: "NR (Para)", color: "Red", tempRange: "-40 to +70 deg C", specificGravity: "1.20 grs/cm3", shoreA: "45 Shore A +/- 5", tensileMpa: "12", elongationPct: "500", abrasion: "150 mm3", thicknesses: "3; 4; 5; 6; 8; 10; 15; 20 mm", widths: "1.5/2 m", lengths: "10 m", features: "Natural rubber quality with good mechanical properties, abrasion resistance, low temperature flexibility and excellent adhesion.", application: paraUse, articleNumbers: codes("ABE45120R", ["3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "15/0", "20/0"]) }),
  p({ productId: "Abrasuper40", group: gp.mining, product: "Red Dragoon", page: 36, quality: "NR (Para)", color: "Red", finish: "One side smooth, other side fabric finish", tempRange: "-40 to +75 deg C", specificGravity: "0.98 grs/cm3", shoreA: "38 Shore A +/- 5", tensileMpa: "25", elongationPct: "900", abrasion: "80 mm3", thicknesses: "2; 3; 4; 5; 6; 8; 10; 30 mm", widths: "1.4 m", lengths: "5/10 m", features: "Premium resistance to abrasion with excellent tensile strength and high elasticity.", application: "Lining of slurry chutes, wear protection, conveyor systems and shooting ranges.", articleNumbers: codes("Abrasuper40", ["2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "30/0"]) }),
  p({ productId: "AB35100R", group: gp.mining, product: "Red Dragoon Eco", page: 37, quality: "NR (Para)", color: "Red", tempRange: "-25 to +70 deg C", specificGravity: "0.98 grs/cm3", shoreA: "35 Shore A +/- 5", tensileMpa: "20", elongationPct: "650", abrasion: "88 mm3", thicknesses: "3; 6; 10; 15; 20 mm", widths: "2 m", lengths: "10 m", features: "Premium resistance to abrasion with excellent tensile strength and high elasticity; bonding layer available.", application: "Lining of slurry chutes, wear protection, conveyor systems and shooting ranges.", articleNumbers: codes("AB35100R", ["3/0", "6/0", "10/0", "15/0", "20/0"]) }),
  p({ productId: "AB40110", group: gp.mining, product: "Centaur Black 40", page: 38, quality: "NR (Para)", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.10 grs/cm3", shoreA: "40 Shore A +/- 5", tensileMpa: "17", elongationPct: "600", abrasion: "120 mm3", thicknesses: "4; 6; 8; 10; 15; 20 mm", widths: "1.4 m", lengths: "10 m", features: "Good resistance to abrasion with high tensile strength and elasticity.", application: "Shot blasting industry, rubber lining for material handling equipment and wear-reduction applications.", articleNumbers: codes("AB40110", ["4/0", "6/0", "8/0", "10/0", "15/0", "20/0"]) }),
  p({ productId: "AB60117", group: gp.mining, product: "Centaur Black (BL)", page: 39, quality: "NR (Para)", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.17 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "13", elongationPct: "450", abrasion: "120 mm3", thicknesses: "3; 4; 5; 6; 8; 10; 12; 15; 20; 25; 30 mm", widths: "1.4/2 m; 2 m on selected thicknesses", lengths: "10 m", features: "Good resistance to abrasion with high tensile strength, elasticity and excellent adhesion; bonding layer available.", application: "Rubber lining for material handling equipment and wear-reduction applications.", articleNumbers: codes("AB60117", ["3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "12/0", "15/0", "20/0", "25/0", "30/0"]) }),
  p({ productId: "NB60135", group: gp.mining, product: "Centaur NBR (BL)", page: 40, quality: "NBR/NR", color: "Black", tempRange: "-25 to +80 deg C", specificGravity: "1.35 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "6", elongationPct: "250", abrasion: "220 mm3", thicknesses: "8; 10 mm", widths: "2 m", lengths: "10 m", features: "Good resistance to oil and abrasion with high tensile strength and elasticity; bonding layer available.", application: "Rubber lining for material handling equipment and wear-reduction applications.", articleNumbers: codes("NB60135", ["8/0", "10/0"]) }),
  p({ productId: "PL60117B", group: gp.mining, product: "Pulley Lagging (BL)", page: 40, quality: "NR (Para)", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.17 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "13", elongationPct: "450", abrasion: "120 mm3", thicknesses: "6.6; 8; 10; 12; 15 mm", widths: "1.5/2 m; 2 m", lengths: "10 m", features: "Diamond or square shaped anti-slip profile with bonding layer.", application: "Reduce belt slippage, wear of the drum and conveyor belt.", articleNumbers: codes("PL60117B", ["6/0", "8/0", "10/0", "12/0", "15/0"]), note: "Catalog table includes mini, square/mini and mini/big design options." }),
  p({ productId: "PL60135NB", group: gp.mining, product: "Pulley Lagging NBR BL", page: 41, quality: "NBR/NR", color: "Black", tempRange: "-25 to +80 deg C", specificGravity: "1.35 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "6", elongationPct: "250", abrasion: "220 mm3", thicknesses: "8; 10 mm", widths: "2 m", lengths: "10 m", features: "Oil and abrasion resistant drum lagging with diamond shaped anti-slip profile with bonding layer.", application: "Reduce belt slippage, wear of the drum and conveyor belt.", articleNumbers: codes("PL60135NB", ["8/0", "10/0"]) }),
  p({ productId: "PL55110B", group: gp.mining, product: "Pulley Lagging BL Food Grade", page: 41, quality: "NR (Para)", color: "Off white, Blue", tempRange: "-30 to +70 deg C", specificGravity: "1.10 grs/cm3", shoreA: "55 Shore A +/- 5", tensileMpa: "15", elongationPct: "600", abrasion: "120 mm3", certification: "FDA and EC1935", thicknesses: "8; 10 mm", widths: "2 m", lengths: "10 m", features: "Food grade drum lagging with good abrasion resistance and diamond shaped anti-slip profile with bonding layer.", application: "Reduce belt slippage, wear of the drum and conveyor belt in the food industry.", articleNumbers: codes("PL55110B", ["8/0", "10/0"]) }),
  p({ productId: "PL45110B", group: gp.mining, product: "Pulley Lagging BL Red", page: 42, quality: "NR (Para)", color: "Red", tempRange: "-30 to +70 deg C", specificGravity: "1.10 grs/cm3", shoreA: "45 Shore A +/- 5", tensileMpa: "15", elongationPct: "600", abrasion: "100 mm3", thicknesses: "8 mm", widths: "2 m", lengths: "10 m", features: "Abrasion resistant diamond shaped anti-slip profile with bonding layer.", application: "Reduce belt slippage, wear of the drum and conveyor belt.", articleNumbers: codes("PL45110B", ["8/0"]) }),
  p({ productId: "PLC70110", group: gp.mining, product: "Ceramic Pulley Lagging BL", page: 42, quality: "NR (Para)", color: "Black", tempRange: "-40 to +70 deg C", specificGravity: "1.10 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "17", elongationPct: "450", abrasion: "100 mm3", thicknesses: "12; 15 mm", widths: "0.5 m", lengths: "5 m", features: "Ceramic equipped drum lagging with excellent abrasion resistance.", application: "Protective surface to improve and extend pulley life in extreme wet and muddy conditions.", articleNumbers: codes("PLC70110", ["12/0", "15/0"]) }),
  p({ productId: "SW64110", group: gp.mining, product: "Triton Sandwich", page: 43, quality: "NR (Para)", color: "Black, Red, Black", tempRange: "-40 to +70 deg C", specificGravity: "1.15 grs/cm3", shoreA: "60/45/60 Shore A +/- 5", tensileMpa: "15", elongationPct: "500", abrasion: "100 mm3", thicknesses: "15; 20; 25 mm", widths: "1.5 m", lengths: "10 m", features: "Unique two colour rubber sheet with good abrasion and tear resistance.", application: "Scraper rubber to clean conveyor belts.", articleNumbers: codes("SW64110", ["15/0", "20/0", "25/0"]) }),
  p({ productId: "AB40100SP", group: gp.mining, product: "Typhon Dustsealing", page: 43, quality: "NR (Para)", color: "Black", tempRange: "-25 to +70 deg C", specificGravity: "0.98 grs/cm3", shoreA: "40 Shore A +/- 5", tensileMpa: "20", elongationPct: "650", abrasion: "90 mm3", thicknesses: "1.5; 2 mm", widths: "1.4 m", lengths: "20 m", features: "Soft black NR sheet with both sides fabric finish.", application: "Dust sealing.", articleNumbers: codes("AB40100SP", ["1,5/0", "2/0"]) }),
  p({ productId: "REPAIR_STRIPS_PATCHES", group: gp.mining, product: "Repair Strips & Patches", page: 44, quality: "NR (Para)", color: "", tempRange: "-40 to +70 deg C", specificGravity: "1.15 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "17", elongationPct: "470", abrasion: "120 mm3", thicknesses: "Patches 2.3/3.7 mm; strips 1.7/2.2/3.2/4.6 mm", widths: "Patch widths 160/260/360 mm; strip widths 50/70/100/150/220/300/400 mm", lengths: "Patches 130/200/270 mm; strips 10 m", features: "Neoprene backed strips and patches to repair conveyor belts; available with and without EP insertion.", application: "Repair of conveyor belts.", articleNumbers: ["PATCH WO 130", "PATCH WO 200", "PATCH WO 270", "PATCH WI 130", "PATCH WI 200", "PATCH WI 270", "STRIP WO 50", "STRIP WO 70", "STRIP WO 100", "STRIP WO 150", "STRIP WO 220", "STRIP WO 300", "STRIP WI 70", "STRIP WI 100", "STRIP WI 150", "STRIP WI 220", "STRIP WI 300", "STRIP WI 400"] }),

  p({ productId: "SB60150F", group: gp.food, product: "Pegasus SBR", page: 47, quality: "NR/SBR", color: "Off white", tempRange: "-30 to +70 deg C", specificGravity: "1.5 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "6", elongationPct: "300", certification: "FDA and EC1935/2004", thicknesses: "1; 2; 3; 4; 5; 6; 8; 10; 15 mm", widths: "1.4 m", lengths: "5/10/20 m", features: "White, odourless and tasteless SBR rubber with excellent mechanical properties.", application: "Food industry.", articleNumbers: codes("SB60150F", ["1/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "15/0"]) }),
  p({ productId: "NB60140F", group: gp.food, product: "Pegasus NBR", page: 48, quality: "NBR", color: "Off white", tempRange: "-35 to +100 deg C", specificGravity: "1.40", shoreA: "60 Shore A +/- 5", tensileMpa: "6", elongationPct: "350", certification: "FDA and EC1935/2004", thicknesses: "1; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/20 m", features: "White, odourless and tasteless NBR rubber, resistant to natural oils and grease.", application: "Food industry.", articleNumbers: codes("NB60140F", ["1/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "EP60130F", group: gp.food, product: "Pegasus EPDM", page: 49, quality: "EPDM", color: "Off white", tempRange: "-25 to +140 deg C", specificGravity: "1.30", shoreA: "60 Shore A +/- 5", tensileMpa: "5", elongationPct: "300", certification: "FDA and EC1935/2004", thicknesses: "1; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10/20 m", features: "White, odourless and tasteless EPDM rubber with a high temperature range.", application: "Food industry.", articleNumbers: codes("EP60130F", ["1/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "NE65145F", group: gp.food, product: "Pegasus CR", page: 50, quality: "CR", color: "Off white", tempRange: "-35 to +120 deg C", specificGravity: "1.45", shoreA: "65 Shore A +/- 5", tensileMpa: "5", elongationPct: "250", certification: "FDA", thicknesses: "1; 2; 3; 5 mm", widths: "1.4 m", lengths: "10/20 m", features: "White, odourless and tasteless CR rubber.", application: "Food industry.", articleNumbers: codes("NE65145F", ["1/0", "2/0", "3/0", "5/0"]) }),
  p({ productId: "SI60120F", group: gp.food, product: "Viking", page: 51, quality: "Silicone", color: "Transparent, Red, Blue, White", tempRange: "-90 to +230 deg C", specificGravity: "1.20", shoreA: "60 Shore A +/- 5", tensileMpa: "6", elongationPct: "200", certification: "FDA and EC1935/2004", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10 mm", widths: "1.2 m", lengths: "5/10 m", features: "Excellent silicone quality with resistance to extreme temperatures.", application: "Food industry, building industry, freezers, boilers, railway industry etc.", articleNumbers: codes("SI60120F", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "SI40120F", group: gp.food, product: "Viking 40", page: 52, quality: "Silicone", color: "Transparent", tempRange: "-90 to +230 deg C", specificGravity: "1.20", shoreA: "40 Shore A +/- 5", tensileMpa: "6", elongationPct: "200", certification: "FDA and EC1935/2004", thicknesses: "2; 3; 4 mm", widths: "1.2 m", lengths: "10 m", features: "Excellent silicone quality with resistance to extreme temperatures.", application: "Food industry, building industry, freezers, boilers, railway industry etc.", articleNumbers: codes("SI40120F", ["2/0", "3/0", "4/0"]) }),

  p({ productId: "FPG72190G", group: gp.fkm, product: "Karibu FKM A", page: 55, quality: "FKM", color: "Black", tempRange: "-30 to +250 deg C", specificGravity: "1.95 grs/cm3", shoreA: "72 Shore A +/- 5", tensileMpa: "5", elongationPct: "200", thicknesses: "1; 1.5; 2; 3; 4; 5; 6; 8; 10; 12 mm", widths: "1.2 m", lengths: "5/10 m", features: "Excellent resistance to oxygen, ozone, weathering, compression at elevated temperatures, solvents, chemicals, hydrocarbons, acids and alkalis.", application: "Gaskets in extreme circumstances.", articleNumbers: codes("FPG72190G", ["1/0", "1,5/0", "2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0", "12/0"]) }),

  p({ productId: "NB65145G", group: gp.various, product: "Kabira Neobenzid", page: 59, quality: "NBR/PVC", color: "Green", tempRange: "-25 to +100 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "6.5", elongationPct: "350", thicknesses: "2; 3; 4; 5; 6; 8; 10 mm", widths: "1.4 m", lengths: "5/10 m", features: "Good resistance to oils and good mechanical properties, no permeability of gases and moderate aging properties.", application: "Petrochemical industries, garages, workshops etc.", articleNumbers: codes("NB65145G", ["2/0", "3/0", "4/0", "5/0", "6/0", "8/0", "10/0"]) }),
  p({ productId: "HN65140", group: gp.various, product: "Serra CSM", page: 59, quality: "CSM", color: "Black", tempRange: "-30 to +140 deg C", specificGravity: "1.40 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "10", elongationPct: "350", thicknesses: "2; 3; 4; 5; 6 mm", widths: "1.4 m", lengths: "10 m", features: "CSM rubber is very resistant to oxygen and ozone, with excellent weathering, abrasion and flame resistance.", application: "", articleNumbers: codes("HN65140", ["2/0", "3/0", "4/0", "5/0", "6/0"]) }),
  p({ productId: "IR60120", group: gp.various, product: "Rama Butyl", page: 60, quality: "IR", color: "Black", tempRange: "-40 to +125 deg C", specificGravity: "1.20 grs/cm3", shoreA: "60 Shore A +/- 5", tensileMpa: "9", elongationPct: "350", thicknesses: "2; 3; 4; 5 mm", widths: "1.4 m", lengths: "10 m", features: "Butyl rubber quality with very low permeability to air and gases, excellent resistance to ozone, oxidation and sunlight.", application: "Gaskets, vibration damping and shock absorption applications.", articleNumbers: codes("IR60120", ["2/0", "3/0", "4/0", "5/0"]) }),

  p({ productId: "SX70150C", group: gp.matting, product: "Cobra", page: 63, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.5 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "3", elongationPct: "200", thicknesses: "3; 4; 5; 6; 8 mm", widths: "1.7/1.8/2 m; 2 m; 1.7/2 m", lengths: "10/12 m", features: "Excellent anti-slip floormat available in 2 meter width.", application: "Car trunks, horse trailers, entrances etc.", articleNumbers: codes("SX70150C", ["3/0", "4/2", "5/0", "6/2", "8/1"]) }),
  p({ productId: "ELS", group: gp.matting, product: "Feniks Insulation Mat", page: 63, quality: "EPDM/SBR", color: "", tempRange: "-25 to +120 deg C", specificGravity: "1.5 grs/cm3", shoreA: "70 Shore A +/- 5", tensileMpa: "5", elongationPct: "250", abrasion: "300 mm3", certification: "IEC CEI 61111", thicknesses: "3; 4.5 mm", widths: "1/1.2 m", lengths: "10 m", features: "Strong and abrasion resistant; electrically insulating according to IEC CEI 61111.", application: "Switch board flooring to prevent electrical shocks.", articleNumbers: ["ELS30000 3/0", "ELS50000 4,5/0"], note: "Proof test voltage lines in catalog: 20/40 kV with article references." }),
  p({ productId: "MR65", group: gp.matting, product: "Saturn Fine Rib", page: 64, quality: "NR/SBR", color: "Black, Grey", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3; 4 mm", widths: "0.2/0.3/0.4/0.5/0.6; 0.7/1/1.2/1.4; 1/1.2 m", lengths: "10 m", features: matBasicFeature, application: matBasicUse, articleNumbers: ["MR65 3/0", "MR65G 3/0", "MR65 4/0"] }),
  p({ productId: "MR65_BASIC", group: gp.matting, product: "Saturn Basic Fine Rib", page: 64, quality: "NR/SBR; NBR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "2.7 mm", widths: "1/1.2; 1/1.2/1.4 m", lengths: "10 m", features: matBasicFeature, application: matBasicUse, articleNumbers: ["MR65SX 2,7/0", "MR65NBR 2,7/0"] }),
  p({ productId: "MN65", group: gp.matting, product: "Saturn Eco Fine Rib", page: 65, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "2.5 mm", widths: "1/1.2 m", lengths: "10 m", features: matBasicFeature, application: matBasicUse, articleNumbers: ["MN65 2,5/0"] }),
  p({ productId: "CR65", group: gp.matting, product: "Athena Broad Rib", page: 65, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3 mm", widths: "1/1.2/1.4/1.6 m", lengths: "10 m", features: matBasicFeature, application: matBasicUse, articleNumbers: ["CR65 3/0"] }),
  p({ productId: "LS65", group: gp.matting, product: "Freva Block Rib", page: 66, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3 mm", widths: "1/1.2 m", lengths: "10 m", features: matBasicFeature, application: matBasicUse, articleNumbers: ["LS65 3/0"] }),
  p({ productId: "BR65", group: gp.matting, product: "Hurricane Wide Rib", page: 66, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "5; 6 mm", widths: "1/1.2; 1.2 m", lengths: "10 m", features: "Excellent non-slip surface with fine cloth impression on the reverse; abrasion resistance helps prevent accidents in all weather conditions.", application: "Multipurpose mat for wet and dry areas, easy to clean with little maintenance.", articleNumbers: ["BR65 5/0", "BR65 6/0"] }),
  p({ productId: "BR65SX", group: gp.matting, product: "Hurricane Eco Wide Rib", page: 67, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.5 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3", elongationPct: "200", thicknesses: "5; 6 mm", widths: "1/1.2 m", lengths: "10 m", features: "Excellent non-slip surface with fine cloth impression on the reverse; abrasion resistance helps prevent accidents in all weather conditions.", application: "Multipurpose mat for wet and dry areas, easy to clean with little maintenance.", articleNumbers: ["BR65SX 5/0", "BR65SX 6/0"] }),
  p({ productId: "TL65", group: gp.matting, product: "Hector Truck Rib", page: 67, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "4.5; 6 mm", widths: "2 m", lengths: "10 m", features: "Deep rib design surface on one side and fine cloth impression pattern on the other side.", application: "Non-skid grip for truck beds to keep cargo in place and reduce sliding damage.", articleNumbers: ["TL65 4,5/0", "TL65 6/0"] }),
  p({ productId: "CS80", group: gp.matting, product: "Vidar Coin Mat", page: 68, quality: "NR/SBR", color: "Black, Grey, Dark Blue, Brown", tempRange: "-30 to +70 deg C", specificGravity: "1.6 grs/cm3", shoreA: "80 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "500 mm3", thicknesses: "3; 4.5 mm", widths: "1/1.2 m", lengths: "10 m", features: "Raised circular design surface on one side and fine cloth impression pattern on the other side.", application: "Indoor applications such as airports, hotels, cabins and exhibitions.", articleNumbers: ["CS80 3/0", "CS80 4,5/0"] }),
  p({ productId: "SD65", group: gp.matting, product: "Sphynx Pyramid Mat", page: 68, quality: "NR/SBR", color: "Black, Grey", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3 mm", widths: "1.2 m", lengths: "10 m", features: "High relief pyramid design on one side and fine cloth impression pattern on the other side.", application: matBasicUse, articleNumbers: ["SD65 3/0"] }),
  p({ productId: "CD65", group: gp.matting, product: "Odin Checker Mat", page: 69, quality: "NR/SBR; NBR/SBR", color: "Black, Grey", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3; 5 mm", widths: "1/1.2/1.4/1.5/1.6; 1.4 m", lengths: "10 m", features: "Continuous checkered field design surface on one side and fine cloth impression pattern on the other side.", application: "Car trunks, horse trailers, entrances etc.", articleNumbers: ["CD65 3/0", "CD65G 3/0", "CD65NBR 3/0", "CD65 5/0"] }),
  p({ productId: "CD65SX", group: gp.matting, product: "Odin Eco Checker Mat", page: 69, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.5 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3", elongationPct: "200", thicknesses: "3 mm", widths: "1/1.2/1.4/1.5/1.6 m", lengths: "10 m", features: "Continuous checkered field design on one side and fine cloth impression pattern on the other side.", application: "Car trunks, horse trailers, entrances etc.", articleNumbers: ["CD65SX 3/0"] }),
  p({ productId: "DCD65", group: gp.matting, product: "Seltos Diamond Checker", page: 70, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3; 5; 6; 8 mm", widths: "1.5/1.8/2/2.2 m; 2 m", lengths: "10 m", features: "Floormat with continuous tear field design on one side and fine cloth impression pattern on the other; available up to 2.2 meter width.", application: "Car trunks, horse trailers, entrances etc.", articleNumbers: ["DCD65 3/0", "DCD65 5/0", "DCD65 6/2", "DCD65 8/1"] }),
  p({ productId: "AM65", group: gp.matting, product: "Amoeba Mat", page: 70, quality: "NR/SBR", color: "Black", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3 mm", widths: "", lengths: "10 m", features: "Anti-slip surface on one side and fine cloth impression pattern on the other side.", application: matBasicUse, articleNumbers: ["AM65 3/0"] }),
  p({ productId: "DO65", group: gp.matting, product: "Zenith Diamond Mat", page: 71, quality: "NR/SBR", color: "Black, Grey", tempRange: "-30 to +70 deg C", specificGravity: "1.45 grs/cm3", shoreA: "65 Shore A +/- 5", tensileMpa: "3.5", elongationPct: "250", abrasion: "300 mm3", thicknesses: "3 mm", widths: "1.2 m", lengths: "10 m", features: "High relief diamond design surface on one side and fine cloth impression pattern on the other side.", application: matBasicUse, articleNumbers: ["DO65 3/0"] }),
  p({ productId: "AV", group: gp.matting, product: "Novibramat", page: 71, quality: "NR/SBR", color: "Black", tempRange: "-20 to +80 deg C", specificGravity: "1.2 grs/cm3", shoreA: "45 or 70 Shore A +/- 5", tensileMpa: "5", elongationPct: "300", thicknesses: "8/10 mm", widths: "0.5 m", lengths: "0.5 m", features: "Anti-vibration rubber mat with rib design at 90 degree angle.", application: "Prevent machine vibration and reduce noise.", articleNumbers: ["AV45 0,5", "AV45 1", "AV70 0,5", "AV70 1"], note: "Catalog states max load 8 kg/cm2." }),
];

function unique(items) {
  return [...new Set(items.filter(Boolean))];
}

function hasAny(text, words) {
  const value = text.toLowerCase();
  return words.some((word) => value.includes(word));
}

function classifyProduct(prod) {
  const text = `${prod.group} ${prod.product} ${prod.quality} ${prod.features} ${prod.application} ${prod.certification} ${prod.note}`.toLowerCase();
  const usage = [];
  const properties = [];
  const recommended = [];
  const basis = [];

  if (prod.group === gp.sbr) {
    usage.push("Üldine tööstus", "Tihendid", "Amortisaatorid", "Konveierid", "Isolatsioon");
    properties.push("Üldotstarbeline", "Löögikindel", "Kulumiskindel", "Mehaaniliselt tugev");
    basis.push("Kataloog: general purpose / industrial rubber");
  }
  if (prod.group === gp.nbr) {
    usage.push("Keemiatööstus", "Garaažid ja töökojad", "Õli ja rasv", "Tihendid");
    properties.push("Õlikindel", "Rasvakindel", "Kütusekindel", "Gaasile vähe läbilaskev", "Mehaaniliselt tugev");
    basis.push("Kataloog: oil and grease resistant NBR");
  }
  if (prod.group === gp.epdm) {
    usage.push("Välitingimused", "Autotööstus", "Päikesepaneelid", "Tihendid");
    properties.push("Ilmastikukindel", "Osoonikindel", "UV-kindel", "Külmakindel", "Keemiakindel");
    basis.push("Kataloog: outdoor/weather/ozone resistant EPDM");
  }
  if (prod.group === gp.cr) {
    usage.push("Offshore", "Tööstus", "Õli ja kemikaalid", "Välitingimused");
    properties.push("Õlikindel", "Osoonikindel", "Ilmastikukindel", "Leegiaeglustav valik", "Mehaaniliselt tugev");
    basis.push("Kataloog: chloroprene with oil, ozone and weather resistance");
  }
  if (prod.group === gp.mining) {
    usage.push("Kaevandus", "Konveierid", "Kulumiskaitse", "Materjalikäitlus");
    properties.push("Kulumiskindel", "Löögikindel", "Rebenemiskindel", "Kõrge tõmbetugevus");
    basis.push("Kataloog: mining/wear/impact/conveyor applications");
  }
  if (prod.group === gp.food || text.includes("food grade") || text.includes("fda") || text.includes("ec1935")) {
    usage.push("Toiduainetööstus", "Toidukontakt", "Tihendid");
    properties.push("Toidukõlblik", "FDA/EC1935", "Lõhnatu/maitsetu valik");
    basis.push("Kataloog: food grade / FDA / EC1935");
  }
  if (prod.group === gp.fkm) {
    usage.push("Ekstreemsed tihendid", "Keemiatööstus", "Kõrge temperatuur", "Süsivesinikud");
    properties.push("Kõrge temperatuuritaluvus", "Keemiakindel", "Õlikindel", "Lahustikindel", "Osoonikindel");
    basis.push("Kataloog: FKM chemical and high-temperature resistance");
  }
  if (prod.group === gp.various) {
    usage.push("Eriotstarbeline tööstus");
    properties.push("Eriotstarbeline");
  }
  if (prod.group === gp.matting) {
    usage.push("Põrandakatted", "Libisemisvastased pinnad", "Tööstuslikud käiguteed");
    properties.push("Libisemisvastane", "Kulumiskindel", "Põrandamatt");
    basis.push("Kataloog: rubber matting / anti-slip use");
  }

  if (hasAny(text, ["abrasion", "wear", "conveyor", "scraper", "slurry", "shot blasting", "material handling"])) {
    usage.push("Konveierid", "Kulumiskaitse");
    properties.push("Kulumiskindel", "Hõõrdumiskindel");
  }
  if (hasAny(text, ["oil", "grease", "gasoline", "bio-diesel", "hydrocarbons", "petrochemical"])) {
    usage.push("Õli ja rasv", "Nafta- ja kütusekeskkond");
    properties.push("Õlikindel", "Kütusekindel");
  }
  if (hasAny(text, ["ozone", "weather", "outdoor", "sunlight", "open air"])) {
    usage.push("Välitingimused");
    properties.push("Ilmastikukindel", "Osoonikindel", "UV-kindel");
  }
  if (hasAny(text, ["potable water", "drinking water"])) {
    usage.push("Joogivesi", "Veetöötlus");
    properties.push("Joogiveega sobiv");
  }
  if (hasAny(text, ["electrically insulating", "electrical shocks", "switch board"])) {
    usage.push("Elektrikilbiruumid", "Elektriisolatsioon");
    properties.push("Elektrit isoleeriv");
  }
  if (hasAny(text, ["vibration", "shock absorption", "anti-vibration"])) {
    usage.push("Vibratsioonisummutus", "Masinaalused");
    properties.push("Vibratsioonisummutav", "Löögisummutav");
  }
  if (hasAny(text, ["truck", "horse trailers", "car trunks", "entrances"])) {
    usage.push("Transport", "Haagised", "Sissepääsud");
  }
  if (hasAny(text, ["freezers", "boilers", "railway", "extreme temperatures"])) {
    usage.push("Külmikud", "Katlad", "Raudtee");
    properties.push("Lai temperatuurivahemik");
  }
  if (hasAny(text, ["flame", "fire"])) {
    properties.push("Leegiaeglustav");
  }
  if (hasAny(text, ["bonding layer"])) {
    properties.push("Liimikihiga valik");
  }

  const snowPlowProducts = new Set([
    "Typhon",
    "Centaur Red (BL)",
    "Centaur Red Eco",
    "Red Dragoon",
    "Red Dragoon Eco",
    "Centaur Black 40",
    "Centaur Black (BL)",
    "Triton Sandwich",
  ]);
  if (snowPlowProducts.has(prod.product)) {
    recommended.push("Lumelükkamine / sahatera kumm");
    usage.push("Lumelükkamine");
    properties.push("Sobib külma ja kulumisega töödeks");
    basis.push("Tuletatud: NR/Para + kulumis-, rebimis- või kaabitsarakendused");
  }
  if (prod.product === "Zenstar" || prod.product === "Black Miami") {
    recommended.push("Üldine sahakumm lihtsamates tingimustes");
    usage.push("Lumelükkamine");
    basis.push("Tuletatud: SBR/NR-SBR üldkumm, löögi- ja kulumiskindlus");
  }
  if (prod.product.includes("Pulley Lagging")) {
    usage.push("Trumlikatted", "Konveierilindi libisemise vähendamine");
    properties.push("Profileeritud pind", "Liimikihiga");
  }
  if (prod.product === "Repair Strips & Patches") {
    usage.push("Konveierilintide remont");
    properties.push("Remondimaterjal", "Neopreen-tagusega");
  }
  if (prod.product === "Typhon Dustsealing") {
    usage.push("Tolmutihendus");
    properties.push("Pehme kumm", "Kangasviimistlusega");
  }
  if (prod.product === "Triton Sandwich") {
    usage.push("Kaabitsad / skreeperid");
    properties.push("Mitmekihiline", "Rebenemiskindel");
  }
  if (prod.product === "Manticore Aqua") {
    usage.push("Joogivesi", "Veetöötlus");
    properties.push("Joogiveega sobiv");
  }
  if (prod.product === "Novibramat") {
    usage.push("Vibratsioonisummutus", "Masinaalused");
    properties.push("Vibratsioonisummutav");
  }

  const allUsage = unique(usage);
  const allProperties = unique(properties);
  const allRecommended = unique(recommended);
  return {
    usageAreas: allUsage,
    propertyTags: allProperties,
    recommendedUses: allRecommended,
    tagBasis: unique(basis),
  };
}

for (const prod of products) {
  Object.assign(prod, classifyProduct(prod));
}

const tagRows = [];
for (const prod of products) {
  for (const tag of prod.usageAreas) {
    tagRows.push([
      prod.productId,
      prod.group,
      prod.product,
      "kasutusvaldkond",
      tag,
      "TULETATUD_VALIKUFILTER",
      sourceText(prod.page),
      `Ei ole eraldi kataloogifakt. Tuletatud tootegrupi, tootekaardi teksti ja materjaliloogika põhjal; kontrolli otseseid kataloogivälju. Alus: ${prod.tagBasis.join("; ")}`,
    ]);
  }
  for (const tag of prod.propertyTags) {
    tagRows.push([
      prod.productId,
      prod.group,
      prod.product,
      "materjali_tunnus",
      tag,
      "TULETATUD_VALIKUFILTER",
      sourceText(prod.page),
      `Ei ole eraldi kataloogifakt. Tuletatud tootegrupi, tootekaardi teksti ja materjaliloogika põhjal; kontrolli otseseid kataloogivälju. Alus: ${prod.tagBasis.join("; ")}`,
    ]);
  }
  for (const tag of prod.recommendedUses) {
    tagRows.push([
      prod.productId,
      prod.group,
      prod.product,
      "soovituslik_kasutus",
      tag,
      "TULETATUD_VALIKUABI",
      sourceText(prod.page),
      `Ei ole kataloogis antud lõplik kasutusgarantii. Vajab ostu/inseneri kontrolli kliendi kasutusolukorra järgi. Alus: ${prod.tagBasis.join("; ")}`,
    ]);
  }
}

const materialMap = new Map();
for (const prod of products) {
  const material = prod.quality || "Määramata";
  if (!materialMap.has(material)) {
    materialMap.set(material, {
      material,
      groups: [],
      products: [],
      usage: [],
      properties: [],
      recommended: [],
      pages: [],
    });
  }
  const row = materialMap.get(material);
  row.groups.push(prod.group);
  row.products.push(prod.product);
  row.usage.push(...prod.usageAreas);
  row.properties.push(...prod.propertyTags);
  row.recommended.push(...prod.recommendedUses);
  row.pages.push(prod.page);
}

const materialRows = [...materialMap.values()]
  .sort((a, b) => a.material.localeCompare(b.material))
  .map((row) => [
    row.material,
    unique(row.groups).join("; "),
    unique(row.products).join("; "),
    unique(row.usage).join("; "),
    unique(row.properties).join("; "),
    unique(row.recommended).join("; "),
    `${Math.min(...row.pages)}-${Math.max(...row.pages)}`,
    catalogSourceFile,
  ]);

const priceByNormArticle = new Map();
for (const row of zenithPriceRows) {
  const key = row.article_norm || normArticle(row.article_nr);
  if (!priceByNormArticle.has(key)) priceByNormArticle.set(key, []);
  priceByNormArticle.get(key).push(row);
}

const priceProductAliases = {
  EP65130: ["EP65135"],
  SI60120F: ["SI60120TP", "SI60120BL", "SI60120BR", "SI60120W"],
  SI40120F: ["SI40120TP"],
  DO65: ["DD65", "DD65G"],
};

function priceMatchesFor(article, productId) {
  const exact = priceByNormArticle.get(normArticle(article));
  if (exact?.length) return exact;

  const articleCompact = compactArticle(article);
  const productCompact = compactArticle(productId);
  const suffix = articleCompact.startsWith(productCompact) ? articleCompact.slice(productCompact.length) : "";
  if (productId === "ELS") {
    const elsAlias = articleCompact.replace(/^ELS30000/, "ELS30000G").replace(/^ELS50000/, "ELS50000G");
    const elsExact = priceByNormArticle.get(normArticle(elsAlias.replace(/(G)(\d)/, "$1 $2")));
    if (elsExact?.length) return elsExact;
    const matches = zenithPriceRows.filter((row) => compactArticle(row.article_nr) === elsAlias);
    if (matches.length) return matches;
  }
  if (productId === "AV") {
    const avBase = articleCompact.startsWith("AV45") ? "AV45" : articleCompact.startsWith("AV70") ? "AV70" : "";
    const avSize = avBase ? articleCompact.slice(avBase.length) : "";
    if (avBase && avSize) {
      const matches = zenithPriceRows.filter((row) => {
        const candidate = compactArticle(row.article_nr);
        return candidate.startsWith(avBase) && candidate.endsWith(avSize);
      });
      if (matches.length) return matches;
    }
  }
  if (!suffix) return [];

  for (const alias of priceProductAliases[productId] ?? []) {
    const aliasCompact = compactArticle(alias);
    const matches = zenithPriceRows.filter((row) => {
      const candidate = compactArticle(row.article_nr);
      return candidate.startsWith(aliasCompact) && candidate.endsWith(suffix);
    });
    if (matches.length) return matches;
  }

  const matches = [];
  for (const row of zenithPriceRows) {
    const candidate = compactArticle(row.article_nr);
    if (productId === "AB45110") {
      if ((candidate.startsWith("AB45110B") || candidate.startsWith("AB45110G")) && candidate.endsWith(suffix)) {
        matches.push(row);
      }
    } else if (candidate.startsWith(productCompact) && candidate.endsWith(suffix)) {
      matches.push(row);
    }
  }
  return matches;
}

function summarizePriceMatches(matches) {
  const gross = matches.map((row) => row.gross_price_eur).filter((value) => typeof value === "number");
  const nett = matches.map((row) => row.nett_price_eur).filter((value) => typeof value === "number");
  const units = unique(matches.map((row) => row.unit));
  const widths = unique(matches.map((row) => row.width_mm).filter((value) => value !== null && value !== undefined).map(String));
  const lengths = unique(matches.map((row) => row.length_mm).filter((value) => value !== null && value !== undefined).map(String));
  return {
    matchCount: matches.length,
    grossMin: gross.length ? Math.min(...gross) : "",
    grossMax: gross.length ? Math.max(...gross) : "",
    nettMin: nett.length ? Math.min(...nett) : "",
    nettMax: nett.length ? Math.max(...nett) : "",
    units: units.join("; "),
    widths: widths.join("; "),
    lengths: lengths.join("; "),
    source: matches.length ? "Pricelist Zenith 2025 v2 (Plastok).xlsx" : "",
  };
}

function productPriceStats(prod) {
  const matches = [];
  for (const article of prod.articleNumbers) {
    matches.push(...priceMatchesFor(article, prod.productId));
  }
  return summarizePriceMatches(matches);
}

for (const prod of products) {
  prod.priceStats = productPriceStats(prod);
}

const productById = new Map(products.map((prod) => [prod.productId, prod]));

function selectionRow({ useCase, level, productId, buyer, engineer, consumer, checks, basis }) {
  const prod = productById.get(productId);
  if (!prod) throw new Error(`Unknown productId in selection helper: ${productId}`);
  return [
    useCase,
    "TULETATUD_VALIKUABI",
    level,
    productId,
    prod.product,
    prod.group,
    buyer,
    engineer,
    consumer,
    checks,
    sourceText(prod.page),
    basis,
    prod.articleNumbers.join("; "),
  ];
}

const selectionRows = [
  selectionRow({
    useCase: "Lumelükkamine / sahatera kumm",
    level: "Parim valik",
    productId: "Abrasuper40",
    buyer: "Vali, kui vaja kulumisele ja löögile vastupidavat premium-kummi.",
    engineer: "NR Para, Shore 38 A, 25 MPa, 900% venivus, abrasioon 80 mm3; sobib elastseks sahateraks.",
    consumer: "Vastupidav ja elastne lahendus raskemaks lumelükkamiseks.",
    checks: "Kontrolli sobiv paksus, kinnitusviis ja kas keskkonnas on õli/kütust.",
    basis: "Tuletatud kataloogi kulumis-, elastsus- ja lining/scraper omaduste põhjal.",
  }),
  selectionRow({
    useCase: "Lumelükkamine / sahatera kumm",
    level: "Parim valik",
    productId: "AB35100R",
    buyer: "Hea valik, kui soovid kulumiskindlat, kuid ökonoomsemat NR Para lahendust.",
    engineer: "NR Para, Shore 35 A, 20 MPa, 650% venivus, abrasioon 88 mm3; bonding layer saadaval.",
    consumer: "Pehmem ja elastne sahakumm külma ning kulumisega töödeks.",
    checks: "Kontrolli bonding layer vajadust ja töötemperatuuri.",
    basis: "Tuletatud kataloogi wear protection / conveyor andmete põhjal.",
  }),
  selectionRow({
    useCase: "Lumelükkamine / sahatera kumm",
    level: "Hea valik",
    productId: "AB60117",
    buyer: "Sobib, kui vaja tugevamat musta kulumiskummi või liimikihiga varianti.",
    engineer: "NR Para, Shore 60 A, 13 MPa, 450% venivus, abrasioon 120 mm3; bonding layer valikud.",
    consumer: "Tugevam must kumm kulumise ja kraapimise vastu.",
    checks: "Kõvem materjal kui Red Dragoon; kontrolli, kas painduvus on piisav.",
    basis: "Tuletatud kataloogi rubber lining / wear reduction info põhjal.",
  }),
  selectionRow({
    useCase: "Lumelükkamine / sahatera kumm",
    level: "Hea valik",
    productId: "SW64110",
    buyer: "Sobib kaabitsa või skreeperi lahendusse, kus kahekihiline ehitus on kasulik.",
    engineer: "NR Para sandwich, 60/45/60 Shore A, 15 MPa, 500%, abrasioon 100 mm3.",
    consumer: "Kaabitsa tüüpi lahendus, kui kumm peab pinda puhastama või kraapima.",
    checks: "Kontrolli, kas sandwich-paksus ja jäikus sobib konkreetsele sahale.",
    basis: "Kataloog: scraper rubber for conveyor belts; lumekasutus on tuletatud.",
  }),
  selectionRow({
    useCase: "Lumelükkamine / sahatera kumm",
    level: "Lihtsam / ajutine valik",
    productId: "SB65130",
    buyer: "Soodsam üldkumm lihtsamasse kasutusse.",
    engineer: "NR/SBR, Shore 65 A, 6 MPa, 350%, abrasioon 250 mm3; mitte premium kulumiskumm.",
    consumer: "Üldotstarbeline kumm kergemaks tööks.",
    checks: "Ei ole esimene valik raskeks sahateraks; kontrolli kulumiskiirust.",
    basis: "Tuletatud kataloogi general purpose / abrasion andmete põhjal.",
  }),

  selectionRow({
    useCase: "Toiduainetööstus / toidukontakt",
    level: "Õli ja rasva korral",
    productId: "NB60140F",
    buyer: "Vali, kui toidukeskkonnas on looduslikke õlisid või rasva.",
    engineer: "Food grade NBR, Shore 60 A, 6 MPa, 350%; FDA ja EC1935/2004.",
    consumer: "Toidukõlblik õli- ja rasvakindlam kumm.",
    checks: "Küsi konkreetse partii sertifikaat ja kontrolli temperatuurivahemik.",
    basis: "Kataloog: food grade NBR, resistant to natural oils and grease.",
  }),
  selectionRow({
    useCase: "Toiduainetööstus / toidukontakt",
    level: "Kõrgem temperatuur",
    productId: "SI60120F",
    buyer: "Vali, kui vaja väga laia temperatuurivahemikku või silikooni.",
    engineer: "Silicone, -90 kuni +230 deg C, Shore 60 A; FDA ja EC1935/2004.",
    consumer: "Kõrge ja madala temperatuuriga toidukeskkondadesse.",
    checks: "Kontrolli mehaanilist koormust, sest silikoon ei ole kulumiskumm.",
    basis: "Kataloog: silicone food quality with extreme temperature resistance.",
  }),
  selectionRow({
    useCase: "Toiduainetööstus / toidukontakt",
    level: "Üldine toidukumm",
    productId: "SB60150F",
    buyer: "SBR-põhine toidukumm üldiseks kasutuseks.",
    engineer: "NR/SBR, Shore 60 A, 6 MPa, 300%; FDA ja EC1935/2004.",
    consumer: "Valge lõhnatu ja maitsetu kumm toidutööstusse.",
    checks: "Õli/rasva korral eelistada NBR Food Grade varianti.",
    basis: "Kataloog: food quality according to FDA and EC1935/2004.",
  }),
  selectionRow({
    useCase: "Toiduainetööstus / toidukontakt",
    level: "Konveieritrumli kate",
    productId: "PL55110B",
    buyer: "Vali toidutööstuse konveieritrumli libisemise vähendamiseks.",
    engineer: "Food grade NR Para pulley lagging, Shore 55 A, FDA ja EC1935.",
    consumer: "Toidutööstuse trumlikate konveierilindi paremaks haardeks.",
    checks: "Kontrolli profiil, bonding layer ja konkreetse toidukontakti nõuded.",
    basis: "Kataloog: food grade drum lagging.",
  }),

  selectionRow({
    useCase: "Õli / rasv / kütusekeskkond",
    level: "Parim tavavalik",
    productId: "NB65130",
    buyer: "Premium NBR õli, rasva ja kütustega kokkupuuteks.",
    engineer: "NBR, Shore 65 A, 10 MPa, 350%; resistant to oil, super, pentane and bio-diesel.",
    consumer: "Õli- ja kütusekindlam kumm tihenditeks ja tööstuskasutuseks.",
    checks: "Välitingimustes kontrolli UV/ozone mõju; EPDM ei sobi õlikeskkonda.",
    basis: "Kataloog: excellent oil and grease resistant NBR.",
  }),
  selectionRow({
    useCase: "Õli / rasv / kütusekeskkond",
    level: "Hea hinna/omaduse valik",
    productId: "NB70140",
    buyer: "Keskmise klassi NBR/SBR õli- ja gaasikeskkondadesse.",
    engineer: "NBR/SBR, Shore 70 A, 5 MPa, 250%; gasoline, oil, propane, natural gas, coal gas.",
    consumer: "Töökindel valik garaaži, töökoja ja keemiatööstuse jaoks.",
    checks: "Kõrgema nõudluse korral võrdle Satyr Premiumiga.",
    basis: "Kataloog: resistant to gasoline, oil, propane and gases.",
  }),
  selectionRow({
    useCase: "Õli / rasv / kütusekeskkond",
    level: "Keemia / süsivesinikud",
    productId: "FPG72190G",
    buyer: "Vali, kui lisaks õlile on vaja keemia- ja kõrget temperatuuritaluvust.",
    engineer: "FKM, -30 kuni +250 deg C, Shore 72 A; hydrocarbons, acids, alkalis, solvents.",
    consumer: "Kõige nõudlikumatesse õli- ja keemiakeskkondadesse.",
    checks: "Kallim erimaterjal; kontrolli konkreetne kemikaal ja kontsentratsioon.",
    basis: "Kataloog: FKM for demanding chemical and hydrocarbon applications.",
  }),
  selectionRow({
    useCase: "Õli / rasv / kütusekeskkond",
    level: "Petrokeemia",
    productId: "NB65145G",
    buyer: "Roheline NBR/PVC variant petrokeemia ja töökojakeskkondadesse.",
    engineer: "NBR/PVC, Shore 65 A, 6.5 MPa, 350%; good oil resistance and gas impermeability.",
    consumer: "Õlide ja gaasidega töökeskkondade erikumm.",
    checks: "Kontrolli temperatuur +100 deg C piiri ja välitingimuste sobivust.",
    basis: "Kataloog: petrochemical industries, garages, workshops.",
  }),

  selectionRow({
    useCase: "Välitingimused / UV / osoon",
    level: "Parim EPDM valik",
    productId: "EP65120",
    buyer: "Premium EPDM välitingimustesse ja kõrgemale temperatuurile.",
    engineer: "EPDM, -40 kuni +130 deg C, Shore 65 A, 8 MPa, 350%; ozone/weather resistant.",
    consumer: "Ilmastiku- ja osoonikindel kumm õuekasutuseks.",
    checks: "Ära kasuta õli- või kütusekeskkonnas ilma lisakontrollita.",
    basis: "Kataloog: excellent weather, cold and ozone resistance.",
  }),
  selectionRow({
    useCase: "Välitingimused / UV / osoon",
    level: "Kõrgem temperatuur",
    productId: "EP70122PC",
    buyer: "Peroxide-cured EPDM, kui vaja kõrgemat temperatuurivahemikku.",
    engineer: "EPDM, -40 kuni +140 deg C, Shore 70 A; chemical and ozone resistance.",
    consumer: "Vastupidav välitingimuste ja kuumema töö vastu.",
    checks: "Õli/rasvaga kontaktis eelista NBR/CR/FKM lahendust.",
    basis: "Kataloog: high temperature peroxid cured EPDM.",
  }),
  selectionRow({
    useCase: "Välitingimused / UV / osoon",
    level: "Õli + ilmastik kompromiss",
    productId: "NE65140",
    buyer: "CR premium, kui on vaja nii ilmastiku kui mõõdukat õlikindlust.",
    engineer: "CR, Shore 65 A, 10 MPa, 400%; heat, ozone, weathering and oil resistance.",
    consumer: "Mitmekülgne kumm õue- ja tööstuskeskkonda.",
    checks: "Toidukontaktiks kasuta Food Grade variante.",
    basis: "Kataloog: CR weathering, ozone and oil resistance.",
  }),

  selectionRow({
    useCase: "Joogivesi / veekontakt",
    level: "Peamine valik",
    productId: "EP70110W",
    buyer: "Vali joogiveega kokkupuutuvate tihendite ja membraanide jaoks.",
    engineer: "EPDM potable water sheet, -40 kuni +120 deg C, Shore 70 A, 11.5 MPa.",
    consumer: "Joogivee kontaktiks mõeldud EPDM kumm.",
    checks: "Küsi vajalik vastavussertifikaat konkreetse projekti standardile.",
    basis: "Kataloog: potable water rubber sheet, certificates on request.",
  }),

  selectionRow({
    useCase: "Konveier / kulumiskaitse / materjalikäitlus",
    level: "Premium kulumiskaitse",
    productId: "Abrasuper40",
    buyer: "Vali rasketesse kulumis- ja löögitingimustesse.",
    engineer: "Abrasioon 80 mm3, 25 MPa, 900%; slurry chutes, wear protection, conveyors.",
    consumer: "Kõrge kulumiskindlusega kumm rasketesse töötingimustesse.",
    checks: "Kontrolli paigaldusviimistlus ja paksus.",
    basis: "Kataloog: lining, wear protection, conveyor systems.",
  }),
  selectionRow({
    useCase: "Konveier / kulumiskaitse / materjalikäitlus",
    level: "Trumlikate",
    productId: "PL60117B",
    buyer: "Vali konveieritrumli libisemise ja kulumise vähendamiseks.",
    engineer: "Profileeritud NR Para lagging bonding layeriga; mini/square/big disainid.",
    consumer: "Parandab konveierilindi haaret ja kaitseb trumlit.",
    checks: "Vali õige profiil, laius ja liimikiht.",
    basis: "Kataloog: pulley lagging to reduce belt slippage and drum wear.",
  }),
  selectionRow({
    useCase: "Konveier / kulumiskaitse / materjalikäitlus",
    level: "Õline konveierikeskkond",
    productId: "PL60135NB",
    buyer: "Vali, kui trumlikate peab olema ka õlikindlam.",
    engineer: "NBR/NR lagging, abrasioon 220 mm3, oil and abrasion resistant.",
    consumer: "Õlikindlam konveieritrumli kate.",
    checks: "Kontrolli õli tüüp ja temperatuur.",
    basis: "Kataloog: oil and abrasion resistant drum lagging.",
  }),
  selectionRow({
    useCase: "Konveier / kulumiskaitse / materjalikäitlus",
    level: "Remont",
    productId: "REPAIR_STRIPS_PATCHES",
    buyer: "Hoia varuna konveierilintide paranduseks.",
    engineer: "Neoprene backed strips and patches; with/without EP insertion.",
    consumer: "Kiireks konveierilindi paranduseks.",
    checks: "Sobita laiused, pikkused ja insertion konkreetse lindiga.",
    basis: "Kataloog: repair of conveyor belts.",
  }),

  selectionRow({
    useCase: "Elektriisolatsioon",
    level: "Peamine valik",
    productId: "ELS",
    buyer: "Vali elektrikilbiruumi või elektriohu vähendamise põrandakatteks.",
    engineer: "EPDM/SBR insulation mat, IEC CEI 61111; kataloogis 20/40 kV proof-test viited.",
    consumer: "Elektrit isoleeriv matt kilbiruumi või tööala jaoks.",
    checks: "Kontrolli nõutav pinge, sertifikaat ja klass enne tellimist.",
    basis: "Kataloog: electrically insulating according to IEC CEI 61111.",
  }),

  selectionRow({
    useCase: "Põrand / libisemisvastane matt",
    level: "Raske üldkasutus",
    productId: "SX70150C",
    buyer: "Cobra sobib autodesse, haagistesse ja sissepääsudesse.",
    engineer: "NR/SBR, Shore 70 A; anti-slip floormat up to 2 m width.",
    consumer: "Tugev libisemisvastane matt autole, treilerile või sissepääsule.",
    checks: "Vali õige insertion ja laius.",
    basis: "Kataloog: car trunks, horse trailers, entrances.",
  }),
  selectionRow({
    useCase: "Põrand / libisemisvastane matt",
    level: "Käigurajad ja tööpinnad",
    productId: "MR65",
    buyer: "Fine rib matt üldiseks põrandaks, käigurajaks või tööpinnaks.",
    engineer: "NR/SBR, Shore 65 A, ribbed surface, cloth impression reverse.",
    consumer: "Lihtne libisemisvastane ribimatt.",
    checks: "Kontrolli värv, laius ja NBR vajadus.",
    basis: "Kataloog: floors, walkways, runners, loading areas, workbenches.",
  }),
  selectionRow({
    useCase: "Põrand / libisemisvastane matt",
    level: "Märg/kuiv ala",
    productId: "BR65",
    buyer: "Wide rib matt märgadesse ja kuivadesse aladesse.",
    engineer: "NR/SBR, anti-slip surface, abrasion resistance in all weather conditions.",
    consumer: "Kergesti puhastatav libisemisvastane matt.",
    checks: "Kontrolli paksus 5/6 mm ja laius.",
    basis: "Kataloog: wet & dry areas, easy to clean.",
  }),
  selectionRow({
    useCase: "Põrand / libisemisvastane matt",
    level: "Lai teemantmuster",
    productId: "DCD65",
    buyer: "Vali, kui vaja kuni 2.2 m laiust diamond checker mustrit.",
    engineer: "NR/SBR, 3/5/6/8 mm, up to 2.2 m width; nylon insertion options.",
    consumer: "Lai teemantmustriga libisemisvastane põrandamatt.",
    checks: "Kontrolli insertion ja soovitud laius.",
    basis: "Kataloog: floormat available up to 2.2 m width.",
  }),

  selectionRow({
    useCase: "Kõrge temperatuur / kemikaalid",
    level: "Keemia ja kuumus",
    productId: "FPG72190G",
    buyer: "FKM kõige nõudlikumasse keemia- ja kuumuskeskkonda.",
    engineer: "FKM kuni +250 deg C; hydrocarbons, acids, alkalis, oxidants, solvents.",
    consumer: "Erikumm äärmuslikesse tingimustesse.",
    checks: "Kontrolli konkreetse kemikaali sobivus tabelist või tootjalt.",
    basis: "Kataloog: most demanding industrial applications.",
  }),
  selectionRow({
    useCase: "Kõrge temperatuur / kemikaalid",
    level: "Toidukeskkond ja temperatuur",
    productId: "SI60120F",
    buyer: "Silikoon toiduainetööstusesse ja väga laia temperatuurivahemikku.",
    engineer: "Silicone -90 kuni +230 deg C; FDA and EC1935/2004.",
    consumer: "Silikoon kuuma, külma ja toidukontakti jaoks.",
    checks: "Ei ole kulumiskummi asendus; kontrolli mehaanilist koormust.",
    basis: "Kataloog: excellent silicone quality with extreme temperature resistance.",
  }),
  selectionRow({
    useCase: "Kõrge temperatuur / kemikaalid",
    level: "EPDM valik",
    productId: "EP70122PC",
    buyer: "Peroxide EPDM kuni +140 deg C välis- ja veekeskkonda.",
    engineer: "EPDM peroxide cured; ozone, chemical and high temperature range.",
    consumer: "Ilmastiku- ja kuumakindlam EPDM.",
    checks: "Õli/kütus ei ole EPDM-i tugevus.",
    basis: "Kataloog: peroxid cured EPDM high temperature range.",
  }),

  selectionRow({
    useCase: "Vibratsioonisummutus / masinaalused",
    level: "Peamine valik",
    productId: "AV",
    buyer: "Vali masinate vibratsiooni ja müra vähendamiseks.",
    engineer: "NR/SBR anti-vibration mat, 45/70 Shore A variandid, max load 8 kg/cm2.",
    consumer: "Vähendab masina vibratsiooni ja müra.",
    checks: "Kontrolli koormus kg/cm2 ja mõõt 0.5 x 0.5 m.",
    basis: "Kataloog: anti-vibration rubber mat.",
  }),
  selectionRow({
    useCase: "Vibratsioonisummutus / masinaalused",
    level: "Täiendav valik tihendiks/amortisaatoriks",
    productId: "IR60120",
    buyer: "Butyl sobib vibratsiooni, tihendite ja gaasiläbilaskvuse vähendamiseks.",
    engineer: "IR, low gas permeability, ozone/oxidation/sunlight resistance, 9 MPa, 350%.",
    consumer: "Butüülkumm vibratsiooni ja tihendamise jaoks.",
    checks: "Kontrolli mehaaniline koormus ja temperatuur +125 deg C.",
    basis: "Kataloog: gaskets, vibration damping and shock absorption.",
  }),
];

const variantRows = [];
for (const prod of products) {
  for (const article of prod.articleNumbers) {
    const thicknessMatch = article.match(/(\d+(?:[,.]\d+)?)\s*\/\d/);
    const priceMatches = priceMatchesFor(article, prod.productId);
    variantRows.push({
      productId: prod.productId,
      group: prod.group,
      product: prod.product,
      article,
      thickness: thicknessMatch ? thicknessMatch[1].replace(",", ".") : "",
      page: prod.page,
      quality: prod.quality,
      color: prod.color,
      sizeInfo: [prod.thicknesses, prod.widths, prod.lengths].filter(Boolean).join(" | "),
      note: prod.note,
      priceStats: summarizePriceMatches(priceMatches),
    });
  }
}

const missingPriceRows = [
  ...products
    .filter((prod) => !prod.priceStats.matchCount)
    .map((prod) => [
      "toode",
      prod.productId,
      prod.product,
      prod.group,
      "",
      prod.articleNumbers.join("; "),
      "Hinnakirja vastet ei leitud artiklikoodi ega konservatiivse alias-sobitusega.",
      sourceText(prod.page),
    ]),
  ...variantRows
    .filter((row) => !row.priceStats.matchCount)
    .map((row) => [
      "artikkel",
      row.productId,
      row.product,
      row.group,
      row.article,
      "",
      "Hinnakirja vastet ei leitud artiklikoodi ega konservatiivse alias-sobitusega.",
      sourceText(row.page),
    ]),
];

const articleAuditRecords = variantRows.map((row) => {
  const pdfArticleFound = pageContains(row.page, row.article);
  const pdfProductNameFound = pageContains(row.page, row.product);
  const pdfArticleOcrVariant = pageContainsOcrVariant(row.page, row.article);
  const pdfProductNameOcrVariant = pageContainsOcrVariant(row.page, row.product);
  const priceFound = row.priceStats.matchCount > 0;
  let status = "KINNITATUD_AUTOMAATSELT";
  if (!pdfArticleFound && !priceFound) {
    status = "KÄSITSI_KONTROLLIDA_PDF_JA_HIND";
  } else if (!pdfArticleFound) {
    status = "KÄSITSI_KONTROLLIDA_PDF_KOOD";
  } else if (!priceFound) {
    status = "KÄSITSI_KONTROLLIDA_HIND";
  }
  if (!pdfArticleFound && pdfArticleOcrVariant && !priceFound) {
    status = "OCR_VARIANT_KÄSITSI_KINNITADA_PDF_JA_HIND";
  } else if (!pdfArticleFound && pdfArticleOcrVariant) {
    status = "OCR_VARIANT_KÄSITSI_KINNITADA_PDF_KOOD";
  }
  if (status === "KINNITATUD_AUTOMAATSELT" && !pdfProductNameFound && pdfProductNameOcrVariant) {
    status = "OCR_VARIANT_KÄSITSI_KINNITADA_TOOTENIMI";
  } else if (status === "KINNITATUD_AUTOMAATSELT" && !pdfProductNameFound) {
    status = "KÄSITSI_KONTROLLIDA_TOOTENIMI";
  }
  return {
    productId: row.productId,
    group: row.group,
    product: row.product,
    page: row.page,
    article: row.article,
    pdfProductNameFound,
    pdfProductNameOcrVariant,
    pdfArticleFound,
    pdfArticleOcrVariant,
    priceFound,
    priceMatchCount: row.priceStats.matchCount,
    status,
    source: sourceText(row.page),
  };
});

const articleAuditByKey = new Map(articleAuditRecords.map((row) => [`${row.productId}||${row.article}`, row]));

const productAuditRecords = products.map((prod) => {
  const articleAudits = articleAuditRecords.filter((row) => row.productId === prod.productId);
  const pdfProductNameFound = pageContains(prod.page, prod.product);
  const pdfProductIdFound = pageContains(prod.page, prod.productId);
  const pdfProductNameOcrVariant = pageContainsOcrVariant(prod.page, prod.product);
  const pdfProductIdOcrVariant = pageContainsOcrVariant(prod.page, prod.productId);
  const pdfArticleFoundCount = articleAudits.filter((row) => row.pdfArticleFound).length;
  const priceArticleFoundCount = articleAudits.filter((row) => row.priceFound).length;
  const missingPdfArticles = articleAudits.filter((row) => !row.pdfArticleFound).map((row) => row.article);
  const missingPriceArticles = articleAudits.filter((row) => !row.priceFound).map((row) => row.article);
  let status = "KINNITATUD_AUTOMAATSELT";
  if (!pdfProductNameFound || missingPdfArticles.length || missingPriceArticles.length) {
    status = "KÄSITSI_KONTROLLIDA";
  }
  if (pdfProductNameFound && !missingPdfArticles.length && missingPriceArticles.length) {
    status = "PDF_KINNITATUD_HIND_KONTROLLIDA";
  }
  if (pdfProductNameFound && missingPdfArticles.length && !missingPriceArticles.length) {
    status = "PDF_TOODE_KINNITATUD_ARTIKLID_KONTROLLIDA";
  }
  if (!pdfProductNameFound && pdfProductNameOcrVariant) {
    status = "OCR_VARIANT_KÄSITSI_KINNITADA_TOOTENIMI";
  }
  return {
    productId: prod.productId,
    group: prod.group,
    product: prod.product,
    page: prod.page,
    pdfProductNameFound,
    pdfProductIdFound,
    pdfProductNameOcrVariant,
    pdfProductIdOcrVariant,
    catalogArticleCount: prod.articleNumbers.length,
    pdfArticleFoundCount,
    priceArticleFoundCount,
    missingPdfArticles,
    missingPriceArticles,
    status,
    source: sourceText(prod.page),
  };
});

const productAuditById = new Map(productAuditRecords.map((row) => [row.productId, row]));
const pdfProductNameFoundCount = productAuditRecords.filter((row) => row.pdfProductNameFound).length;
const pdfArticleFoundCount = articleAuditRecords.filter((row) => row.pdfArticleFound).length;
const pdfArticleMissingCount = articleAuditRecords.length - pdfArticleFoundCount;
const priceArticleFoundCount = articleAuditRecords.filter((row) => row.priceFound).length;
const priceArticleMissingCount = articleAuditRecords.length - priceArticleFoundCount;

const productAuditRows = productAuditRecords.map((row) => [
  row.productId,
  row.group,
  row.product,
  row.page,
  row.pdfProductNameFound ? "JAH" : "EI",
  row.pdfProductNameOcrVariant ? "JAH" : "EI",
  row.pdfProductIdFound ? "JAH" : "EI",
  row.pdfProductIdOcrVariant ? "JAH" : "EI",
  row.catalogArticleCount,
  row.pdfArticleFoundCount,
  row.catalogArticleCount - row.pdfArticleFoundCount,
  row.priceArticleFoundCount,
  row.catalogArticleCount - row.priceArticleFoundCount,
  row.missingPdfArticles.join("; "),
  row.missingPriceArticles.join("; "),
  row.status,
  row.pdfProductNameOcrVariant || row.pdfProductIdOcrVariant ? "Täpne PDF tekstikihi vaste puudus, kuid OCR-lähedane variant leiti; kinnita visuaalselt PDF-ist." : "",
  row.source,
]);

const articleAuditRows = articleAuditRecords.map((row) => [
  row.productId,
  row.group,
  row.product,
  row.page,
  row.article,
  row.pdfProductNameFound ? "JAH" : "EI",
  row.pdfProductNameOcrVariant ? "JAH" : "EI",
  row.pdfArticleFound ? "JAH" : "EI",
  row.pdfArticleOcrVariant ? "JAH" : "EI",
  row.priceFound ? "JAH" : "EI",
  row.priceMatchCount,
  row.status,
  row.pdfArticleOcrVariant || row.pdfProductNameOcrVariant ? "Täpne PDF tekstikihi vaste puudus, kuid OCR-lähedane variant leiti; kinnita visuaalselt PDF-ist." : "",
  row.source,
]);

const manualControlRows = [
  ...productAuditRecords
    .filter((row) => !row.pdfProductNameFound)
    .map((row) => [
      "Tootenimi",
      row.productId,
      row.product,
      "",
      sourceText(row.page),
      "PDF tekstikihis ei leitud tootenime täpset vastet.",
      row.pdfProductNameOcrVariant ? "OCR-lähedane variant leiti; kinnita visuaalselt PDF-ist." : "Kontrolli visuaalselt PDF-ist.",
      row.status,
    ]),
  ...articleAuditRecords
    .filter((row) => !row.pdfArticleFound || !row.priceFound)
    .map((row) => [
      "Artikkel",
      row.productId,
      row.product,
      row.article,
      sourceText(row.page),
      [!row.pdfArticleFound ? "PDF tekstikihis ei leitud artiklikoodi täpset vastet" : "", !row.priceFound ? "Zenith hinnakirjas ei leitud artiklikoodi vastet" : ""].filter(Boolean).join("; "),
      row.pdfArticleOcrVariant ? "OCR-lähedane artiklikoodi variant leiti; kinnita visuaalselt PDF-ist." : "Kontrolli allikast käsitsi.",
      row.status,
    ]),
];

function sourceControlRows() {
  return [
  ["Zenith kataloog PDF", catalogSourceFile, "PDF lehti", sourceExtract.catalog_page_count ?? 0, 72, (sourceExtract.catalog_page_count ?? 0) === 72 ? "OK" : "KONTROLLIDA"],
  ["Zenith kataloog PDF", catalogSourceFile, "Tootenime vaste PDF tekstikihis", pdfProductNameFoundCount, products.length, pdfProductNameFoundCount === products.length ? "OK" : "KÄSITSI_KONTROLLIDA"],
  ["Zenith kataloog PDF", catalogSourceFile, "Artiklikoodi täpne vaste PDF tekstikihis", pdfArticleFoundCount, variantRows.length, pdfArticleFoundCount === variantRows.length ? "OK" : "KÄSITSI_KONTROLLIDA"],
  ["Zenith hinnakiri", "Pricelist Zenith 2025 v2 (Plastok).xlsx", "ALL lehe imporditud hinnaread", zenithPriceRows.length, "lähtefaili mitte-tühjad read", zenithPriceRows.length ? "OK" : "KONTROLLIDA"],
  ["Zenith hinnakiri", "Pricelist Zenith 2025 v2 (Plastok).xlsx", "Artiklikoodi vaste hinnakirjas", priceArticleFoundCount, variantRows.length, priceArticleFoundCount === variantRows.length ? "OK" : "HINNAVASTETA_READ_LEHEL_14"],
  ["Plastok DATABASE", "DATABASE.xlsm", "DATABASE lehe imporditud read", plastokRows.length, "lähtefaili mitte-tühjad read", plastokRows.length ? "OK" : "KONTROLLIDA"],
  ["Käsitsi kontroll", "21_KÄSITSI_KONTROLL", "Kontrolli vajavad read", manualControlRows.length, "0 oleks täielikult automaatselt kinnitatud", manualControlRows.length ? "KÄSITSI_KONTROLLIDA" : "OK"],
  ["Tuletatud väljad", "build_zenith_catalog_database.mjs", "Kasutusfiltrid / valikuabi / Plastok normaliseerimine", tagRows.length + selectionRows.length + plastokMaterialRows.length + plastokFirstDecisionRows.length, "ei ole kataloogifaktid", "MÄRGITUD_TULETATUD"],
  ];
}

const plastokFamilyMap = new Map();
for (const row of plastokRows) {
  const family = normalizePlastokFamily(row);
  const profile = plastokFamilyProfiles[family] ?? plastokFamilyProfiles["Määramata / kontrollida"];
  if (!plastokFamilyMap.has(family)) {
    plastokFamilyMap.set(family, {
      family,
      profile,
      rows: 0,
      articles: new Set(),
      categories: new Set(),
      groups: new Set(),
      materials: new Set(),
      units: new Set(),
      quantityRows: 0,
      grossPrices: [],
      nettPrices: [],
      sampleArticles: [],
    });
  }
  const item = plastokFamilyMap.get(family);
  item.rows += 1;
  if (row["Article nr."]) item.articles.add(String(row["Article nr."]));
  if (row["Category"]) item.categories.add(String(row["Category"]));
  if (row["Group"]) item.groups.add(String(row["Group"]));
  if (row["Material"]) item.materials.add(String(row["Material"]));
  if (row["Unit"]) item.units.add(String(row["Unit"]));
  if (typeof row["Quantity"] === "number" && row["Quantity"] > 0) item.quantityRows += 1;
  if (typeof row["Gross price, EUR"] === "number" && row["Gross price, EUR"] > 0) item.grossPrices.push(row["Gross price, EUR"]);
  if (typeof row["Nett price, EUR"] === "number" && row["Nett price, EUR"] > 0) item.nettPrices.push(row["Nett price, EUR"]);
  if (item.sampleArticles.length < 8 && row["Article nr."] && row["Name"]) {
    item.sampleArticles.push(`${row["Article nr."]} - ${row["Name"]}`);
  }
}

function sortedSetText(set, limit = 18) {
  const values = [...set].filter(Boolean).sort();
  return values.slice(0, limit).join("; ") + (values.length > limit ? `; ... (+${values.length - limit})` : "");
}

function familyStats(family) {
  return plastokFamilyMap.get(family);
}

function familyBrief(family) {
  const item = familyStats(family);
  if (!item) return `${family}: Plastoki andmebaasis otsene grupp puudub.`;
  return `${family}: ${item.articles.size} artiklit, ${item.quantityRows} kogusega rida, bruto ${fmtNum(Math.min(...item.grossPrices))}-${fmtNum(Math.max(...item.grossPrices))} EUR`;
}

const plastokMaterialRows = [...plastokFamilyMap.values()]
  .sort((a, b) => b.articles.size - a.articles.size)
  .map((item) => [
    item.family,
    "TULETATUD_NORMALISEERIMINE",
    item.profile.role,
    item.profile.uses,
    item.profile.strengths,
    item.profile.cautions,
    item.rows,
    item.articles.size,
    item.quantityRows,
    item.grossPrices.length ? fmtNum(Math.min(...item.grossPrices)) : "",
    item.grossPrices.length ? fmtNum(Math.max(...item.grossPrices)) : "",
    item.nettPrices.length ? fmtNum(Math.min(...item.nettPrices)) : "",
    item.nettPrices.length ? fmtNum(Math.max(...item.nettPrices)) : "",
    sortedSetText(item.units),
    sortedSetText(item.categories),
    sortedSetText(item.groups),
    sortedSetText(item.materials, 24),
    item.sampleArticles.join("; "),
    item.profile.zenithTrigger,
    "DATABASE.xlsm",
  ]);

function decisionRow({ useCase, plastokFamilies, ownEnough, zenithWhen, zenithProducts, checks, risk }) {
  return [
    useCase,
    "TULETATUD_OTSUSTUGI",
    plastokFamilies.join("; "),
    plastokFamilies.map((family) => familyBrief(family)).join(" | "),
    ownEnough,
    zenithWhen,
    zenithProducts,
    checks,
    risk,
    "DATABASE.xlsm + Zenith catalogue 2020_0.pdf + Pricelist Zenith 2025 v2 (Plastok).xlsx",
  ];
}

const plastokFirstDecisionRows = [
  decisionRow({
    useCase: "Lumelükkamine / sahatera kumm",
    plastokFamilies: ["SBR/NR üld- ja kulumiskumm", "PUR / PU"],
    ownEnough: "Plastoki oma kulumiskumm või sahakumm sobib, kui paksus, laius, kõvadus ja kogus on olemas ning töö ei nõua konkreetset Zenith kvaliteeti.",
    zenithWhen: "Telli Zenithist juurde, kui oma kulumiskummi mõõt/kogus ei sobi või vaja Red Dragoon/Typhon tüüpi kontrollitud tarnijakvaliteeti.",
    zenithProducts: "Red Dragoon; Red Dragoon Eco; Typhon; Centaur Black; Triton Sandwich",
    checks: "Paksus, Shore A, laius, kinnitusviis, töötemperatuur, kulumiskoormus, õli/kütuse kontakt.",
    risk: "Liiga pehme või liiga kõva materjal võib sahal kiiresti kuluda või halvasti töötada.",
  }),
  decisionRow({
    useCase: "Toiduainetööstus / toidukontakt",
    plastokFamilies: ["Silikoon", "EPDM kumm", "SBR/NR üld- ja kulumiskumm", "NBR kumm", "PE / UHMW plastid"],
    ownEnough: "Kasuta Plastoki oma FDA/toidukontakti materjali, kui sertifikaat, mõõt ja temperatuur sobivad.",
    zenithWhen: "Telli Zenithist juurde, kui vaja Pegasus/Viking Food Grade seeria kindlat mõõtu või sertifikaati.",
    zenithProducts: "Pegasus SBR; Pegasus NBR; Pegasus EPDM; Pegasus CR; Viking; Viking 40; Pulley Lagging BL Food Grade",
    checks: "FDA/EC1935 sertifikaat, õli/rasv, temperatuur, puhastuskeemia, värv, kõvadus.",
    risk: "Toidukontakti puhul ei piisa materjalinimest; partii/sertifikaat tuleb kinnitada.",
  }),
  decisionRow({
    useCase: "Õli / rasv / kütusekeskkond",
    plastokFamilies: ["NBR kumm", "CR / neopreen", "FKM / Viton", "PTFE / tihendiplaadid"],
    ownEnough: "Plastoki NBR/CR/FKM või tihendimaterjal sobib, kui kemikaal, temperatuur ja mõõt on kinnitatud.",
    zenithWhen: "Telli Zenithist juurde, kui oma NBR/FKM mõõtu ei ole või klient vajab Zenith Satyr/Karibu artiklit.",
    zenithProducts: "Satyr Premium; Satyr Medium; Kabira Neobenzid; Karibu FKM A; Wendigo Premium",
    checks: "Õli tüüp, kütus, temperatuur, UV, tihendi surve, mehaaniline koormus.",
    risk: "EPDM ei ole õli/kütuse jaoks sobiv vaikimisi valik.",
  }),
  decisionRow({
    useCase: "Välitingimused / UV / osoon",
    plastokFamilies: ["EPDM kumm", "CR / neopreen"],
    ownEnough: "Plastoki EPDM/CR sobib, kui mõõt, kõvadus ja kogus on olemas ning õlikontakt puudub või on kontrollitud.",
    zenithWhen: "Telli Zenithist juurde, kui vaja Manticore/Wendigo seeria kindlat tarnijakoodi või Plastoki mõõt puudub.",
    zenithProducts: "Manticore Premium; Manticore Peroxid; Manticore Eco; Wendigo Premium",
    checks: "UV/osoon, temperatuur, õli/rasv, väliskeskkonna kemikaalid, sertifikaat.",
    risk: "NBR võib välitingimustes vananeda kiiremini; EPDM ei sobi õlikeskkonda.",
  }),
  decisionRow({
    useCase: "Joogivesi / veekontakt",
    plastokFamilies: ["EPDM kumm", "PTFE / tihendiplaadid"],
    ownEnough: "Kasuta Plastoki materjali ainult siis, kui joogivee/veekontakti nõutud sertifikaat on olemas.",
    zenithWhen: "Telli Zenith Manticore Aqua, kui vaja kataloogis joogiveele mõeldud EPDM-i ja sertifikaat on saadav/kinnitatud.",
    zenithProducts: "Manticore Aqua",
    checks: "Joogivee standard, partii sertifikaat, temperatuur, puhastuskeemia, membraani/tihendi töötingimus.",
    risk: "Tavaline EPDM ei tähenda automaatselt joogiveesobivust.",
  }),
  decisionRow({
    useCase: "Konveier / kulumiskaitse / materjalikäitlus",
    plastokFamilies: ["SBR/NR üld- ja kulumiskumm", "PUR / PU", "PE / UHMW plastid"],
    ownEnough: "Plastoki kulumiskumm, PU või PE1000 sobib, kui funktsioon on kulumisdetail/juhik ja mõõt on olemas.",
    zenithWhen: "Telli Zenithist juurde, kui vaja kaevandus-/konveierikummi, trumlikatet, repair strippe või kindlat Mining & Para materjali.",
    zenithProducts: "Red Dragoon; Centaur Black; Pulley Lagging; Repair Strips & Patches; Typhon Dustsealing",
    checks: "Kas vaja elastset kummi või jäika plastikut, abrasioon, löök, paksus, liimikiht, profiil.",
    risk: "Plast ja kumm võivad sama probleemi lahendada erinevalt; funktsioon peab olema selge.",
  }),
  decisionRow({
    useCase: "Elektriisolatsioon",
    plastokFamilies: ["Elektriisolatsiooni matid", "Kummimatid"],
    ownEnough: "Plastoki dielektriline matt sobib, kui pingeklass, sertifikaat, paksus ja mõõt vastavad nõudele.",
    zenithWhen: "Telli Zenith Feniks, kui oma dielektrilise mati klass/mõõt ei sobi või vaja Zenith artiklit.",
    zenithProducts: "Feniks Insulation Mat",
    checks: "Pingeklass, standard, sertifikaat, paksus, libisemisvastasus, tööala suurus.",
    risk: "Tavaline kummimatt ei ole automaatselt elektriisolatsiooni matt.",
  }),
  decisionRow({
    useCase: "Põrand / libisemisvastane matt",
    plastokFamilies: ["Kummimatid"],
    ownEnough: "Plastoki matt sobib, kui muster, laius, paksus, värv ja kogus on olemas.",
    zenithWhen: "Telli Zenith matting seeriast juurde, kui vaja konkreetset Cobra/Saturn/Hurricane/Odin/Seltos mustrit või mõõtu.",
    zenithProducts: "Cobra; Saturn Fine Rib; Hurricane Wide Rib; Odin Checker; Seltos Diamond Checker; Novibramat",
    checks: "Muster, puhastatavus, libisemisvastasus, õli/rasv, välitingimus, elektriisolatsiooni nõue.",
    risk: "Muster ja pind on kasutajakogemuse jaoks sama tähtsad kui materjal.",
  }),
  decisionRow({
    useCase: "Kõrge temperatuur / kemikaalid",
    plastokFamilies: ["Silikoon", "FKM / Viton", "PTFE / tihendiplaadid", "EPDM kumm"],
    ownEnough: "Plastoki silikoon/FKM/PTFE sobib, kui temperatuur, kemikaal ja mehaaniline koormus on kinnitatud.",
    zenithWhen: "Telli Zenithist juurde, kui oma mõõt puudub või vaja Karibu FKM / Viking / Manticore Peroxid seeriat.",
    zenithProducts: "Karibu FKM A; Viking; Viking 40; Manticore Peroxid",
    checks: "Kemikaal, kontsentratsioon, temperatuur, toidukontakt, mehaaniline koormus, puhastus.",
    risk: "Silikoon ei ole kulumiskumm; FKM on kallis ja kemikaal peab sobima.",
  }),
  decisionRow({
    useCase: "Vibratsioonisummutus / masinaalused",
    plastokFamilies: ["Vibratsioonimaterjalid", "Butüül / IR", "Kummimatid"],
    ownEnough: "Plastoki vibratsioonimaterjal sobib, kui koormus, paksus ja deformatsioon on sobivad.",
    zenithWhen: "Telli Zenith Novibramat või Rama Butyl, kui oma materjali koormus/mõõt ei sobi.",
    zenithProducts: "Novibramat; Rama Butyl",
    checks: "Koormus kg/cm2, vibratsiooni sagedus, paksus, pind, õli/kemikaal, välitingimus.",
    risk: "Vale koormusvahemik võib summutuse ära rikkuda või materjali kiiresti deformeerida.",
  }),
];

function sheetRange(rowCount, colCount) {
  const letters = [];
  let n = colCount;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters.unshift(String.fromCharCode(65 + rem));
    n = Math.floor((n - 1) / 26);
  }
  return `A1:${letters.join("")}${rowCount}`;
}

function addTableSheet(workbook, name, headers, rows, widths, tableName) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const values = [headers, ...rows];
  sheet.getRangeByIndexes(0, 0, values.length, headers.length).values = values;
  const used = sheet.getRangeByIndexes(0, 0, values.length, headers.length);
  used.format.wrapText = true;
  used.format.verticalAlignment = "top";
  const header = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  header.format = {
    fill: "#174E4A",
    font: { bold: true, color: "#FFFFFF" },
  };
  header.format.rowHeightPx = 34;
  used.format.borders = { preset: "all", style: "thin", color: "#D7DEE2" };
  const table = sheet.tables.add(sheetRange(values.length, headers.length), true, tableName);
  table.style = "TableStyleMedium2";
  sheet.freezePanes.freezeRows(1);
  widths.forEach((px, idx) => {
    sheet.getRangeByIndexes(0, idx, values.length, 1).format.columnWidthPx = px;
  });
  return sheet;
}

function groupCount(group) {
  return products.filter((prod) => prod.group === group).length;
}

function articleCount(group) {
  return variantRows.filter((row) => row.group === group).length;
}

function addSummary(workbook) {
  const sheet = workbook.worksheets.add("00_KOKKUVOTE");
  sheet.showGridLines = false;
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").values = [["Plastok-Zenith materjali- ja valikuabi kokkuvõte"]];
  sheet.getRange("A1").format = {
    fill: "#123C42",
    font: { bold: true, color: "#FFFFFF", size: 16 },
  };
  sheet.getRange("A3:B24").values = [
    ["Allikas", catalogSourceFile],
    ["Kataloogi lehti", 72],
    ["PDF lehti tekstikihis", sourceExtract.catalog_page_count ?? 0],
    ["Tootegruppe", groupDescriptions.length],
    ["Master-tooteid", products.length],
    ["Artikli-/mõõduridu", variantRows.length],
    ["PDF tootenimed leitud", `${pdfProductNameFoundCount}/${products.length}`],
    ["PDF artiklikoodid leitud", `${pdfArticleFoundCount}/${variantRows.length}`],
    ["PDF artiklid käsitsi kontrollida", pdfArticleMissingCount],
    ["Hinnakirja artiklid vastega", `${priceArticleFoundCount}/${variantRows.length}`],
    ["Käsitsi kontrolli read", manualControlRows.length],
    ["Toote-tunnuse ridu", tagRows.length],
    ["Materjali/kvaliteedi gruppe", materialRows.length],
    ["Valikuabi ridu", selectionRows.length],
    ["Zenith hinnakirja ridu", zenithPriceRows.length],
    ["Zenith hinnakoondi ridu", zenithPriceSummaryRows.length],
    ["Plastok DB ridu", plastokRows.length],
    ["Plastok normaliseeritud gruppe", plastokMaterialRows.length],
    ["Plastok-first otsuseridu", plastokFirstDecisionRows.length],
    ["Hinnavasteta tooteid", products.filter((prod) => !prod.priceStats.matchCount).length],
    ["Hinnavasteta artikliridu", variantRows.filter((row) => !row.priceStats.matchCount).length],
    ["Valmimise kuupäev", "2026-06-09"],
  ];
  sheet.getRange("A26:D26").values = [["Tootegrupp", "Kataloogi lehed", "Tooteid", "Artikli-/mõõduridu"]];
  const summaryRows = groupDescriptions.map(([group, , pages]) => [group, pages, groupCount(group), articleCount(group)]);
  sheet.getRangeByIndexes(26, 0, summaryRows.length, 4).values = summaryRows;
  sheet.getRange("A26:D26").format = { fill: "#174E4A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRangeByIndexes(25, 0, summaryRows.length + 1, 4).format.borders = { preset: "all", style: "thin", color: "#D7DEE2" };
  sheet.getRange("A37:F48").values = [
    ["Kontrollimärkused", "", "", "", "", ""],
    [`- ${businessRule}`, "", "", "", "", ""],
    ["- Oletusi ei käsitleta faktidena: otsesed kataloogi/hinnakirja/Plastok DB väljad on eraldi, tuletatud väljad on märgitud TULETATUD_* staatusega.", "", "", "", "", ""],
    ["- PDF tekstikihi kontroll kinnitab tootenime või artiklikoodi ainult siis, kui vastav tekst/kood leiti samalt viidatud lehelt.", "", "", "", "", ""],
    ["- Kui PDF tekstikihi vaste puudub, jääb rida auditilehele käsitsi kontrollimiseks; seda ei loeta automaatselt veaks ega kinnitatud faktiks.", "", "", "", "", ""],
    ["- Zenith kataloog on tarnijakataloog: seda kasutatakse juurde tellimise ja sobivuse kontrolli allikana, mitte Plastoki oma laobaasi asendusena.", "", "", "", "", ""],
    ["- Toodete tase on koostatud Zenith kataloogi tootekaartide põhjal; artikli-/mõõduread on eraldi lehel.", "", "", "", "", ""],
    ["- PDF-i OCR-is olid mõned nimed/koodid moonutatud; failis on need korrigeeritud loetavaks ja kontrollitavaks.", "", "", "", "", ""],
    ["- Vana põhja toode 'Commercial Silicone' ei olnud 2020 kataloogis selge tootekaardina, seega seda ei lisatud kataloogibaasi master-tootena.", "", "", "", "", ""],
    ["- Kasutusvaldkonnad ja tunnused on lisatud filtritavate veergudena; tuletatud soovitused on eraldi markeeritud.", "", "", "", "", ""],
    ["- Hinnad pärinevad Zenith 2025 hinnakirjast Plastokile; hinnavasteta read on eraldi kontrolllehel.", "", "", "", "", ""],
    ["- Plastok DATABASE.xlsm on esmane oma materjalide andmebaas ja seda kasutatakse enne Zenithist tellimise otsust.", "", "", "", "", ""],
  ];
  for (let row = 37; row <= 48; row += 1) {
    sheet.getRange(`A${row}:F${row}`).merge();
  }
  sheet.getRange("A37").format = { fill: "#E7F0EE", font: { bold: true, color: "#123C42" } };
  sheet.getRange("A38:F48").format.wrapText = true;
  [210, 130, 90, 130, 90, 90].forEach((px, idx) => {
    sheet.getRangeByIndexes(0, idx, 51, 1).format.columnWidthPx = px;
  });
  sheet.getRange("A3:B24").format.borders = { preset: "all", style: "thin", color: "#D7DEE2" };
  sheet.getRange("A3:A24").format = { fill: "#E7F0EE", font: { bold: true, color: "#123C42" } };
}

const workbook = Workbook.create();
addSummary(workbook);

addTableSheet(
  workbook,
  "01_TOOTED",
  [
    "product_id",
    "tootegrupp",
    "toode",
    "allikas",
    "kataloogi_leht",
    "andmete_kontrolli_staatus",
    "pdf_tootenimi_leitud",
    "pdf_artiklikoode_leitud",
    "pdf_artiklikoode_kontrollida",
    "tuletatud_kasutusfiltrid",
    "tuletatud_soovituslikud_kasutused",
    "tuletatud_materjali_tunnused",
    "tuletuse_alus",
    "hinnakirja_vasteid",
    "bruto_min_eur",
    "bruto_max_eur",
    "netto_min_eur",
    "netto_max_eur",
    "hinna_ühik",
    "hinna_allikas",
    "kvaliteet_materjal",
    "värv",
    "temperatuur",
    "erikaal",
    "shore_a",
    "tombetugevus_mpa",
    "katkevenivus_pct",
    "abrasioon",
    "sertifikaat",
    "viimistlus",
    "paksused_mm",
    "laiused_m",
    "pikkused_m",
    "omadused",
    "kasutus",
    "artiklid",
    "märkus",
  ],
  products.map((prod) => {
    const audit = productAuditById.get(prod.productId);
    return [
      prod.productId,
      prod.group,
      prod.product,
      sourceText(prod.page),
      prod.page,
      audit?.status ?? "KONTROLL PUUDUB",
      audit?.pdfProductNameFound ? "JAH" : "EI",
      audit ? `${audit.pdfArticleFoundCount}/${audit.catalogArticleCount}` : "",
      audit?.missingPdfArticles.join("; ") ?? "",
      prod.usageAreas.join("; "),
      prod.recommendedUses.join("; "),
      prod.propertyTags.join("; "),
      prod.tagBasis.join("; "),
      prod.priceStats.matchCount,
      fmtNum(prod.priceStats.grossMin),
      fmtNum(prod.priceStats.grossMax),
      fmtNum(prod.priceStats.nettMin),
      fmtNum(prod.priceStats.nettMax),
      prod.priceStats.units,
      prod.priceStats.source,
      prod.quality,
      prod.color,
      prod.tempRange,
      prod.specificGravity,
      prod.shoreA,
      prod.tensileMpa,
      prod.elongationPct,
      prod.abrasion,
      prod.certification,
      prod.finish,
      prod.thicknesses,
      prod.widths,
      prod.lengths,
      prod.features,
      prod.application,
      prod.articleNumbers.join("; "),
      prod.note,
    ];
  }),
  [115, 160, 180, 210, 82, 180, 95, 125, 360, 360, 260, 360, 360, 95, 95, 95, 95, 95, 95, 210, 120, 130, 130, 110, 130, 95, 95, 95, 150, 170, 180, 160, 140, 360, 320, 360, 260],
  "TootedTable",
);

addTableSheet(
  workbook,
  "02_ARTIKLID_MOODUD",
  ["product_id", "tootegrupp", "toode", "article_nr", "andmete_kontrolli_staatus", "pdf_artiklikood_leitud", "allikas", "tuletatud_kasutusfiltrid", "tuletatud_soovituslikud_kasutused", "tuletatud_materjali_tunnused", "hinnakirja_vasteid", "bruto_min_eur", "bruto_max_eur", "netto_min_eur", "netto_max_eur", "hinna_ühik", "hinna_laiused_mm", "hinna_pikkused_mm", "hinna_allikas", "paksus_koodist_mm", "kataloogi_leht", "kvaliteet_materjal", "värv", "mõõdude_kokkuvõte", "märkus"],
  variantRows.map((row) => {
    const prod = products.find((item) => item.productId === row.productId);
    const audit = articleAuditByKey.get(`${row.productId}||${row.article}`);
    return [
      row.productId,
      row.group,
      row.product,
      row.article,
      audit?.status ?? "KONTROLL PUUDUB",
      audit?.pdfArticleFound ? "JAH" : "EI",
      sourceText(row.page),
      prod?.usageAreas.join("; ") ?? "",
      prod?.recommendedUses.join("; ") ?? "",
      prod?.propertyTags.join("; ") ?? "",
      row.priceStats.matchCount,
      fmtNum(row.priceStats.grossMin),
      fmtNum(row.priceStats.grossMax),
      fmtNum(row.priceStats.nettMin),
      fmtNum(row.priceStats.nettMax),
      row.priceStats.units,
      row.priceStats.widths,
      row.priceStats.lengths,
      row.priceStats.source,
      row.thickness,
      row.page,
      row.quality,
      row.color,
      row.sizeInfo,
      row.note,
    ];
  }),
  [115, 155, 180, 150, 190, 110, 210, 340, 250, 340, 95, 95, 95, 95, 95, 95, 150, 150, 210, 105, 82, 120, 130, 390, 260],
  "ArtiklidTable",
);

addTableSheet(
  workbook,
  "03_GRUPID",
  ["tootegrupp", "kataloogi_lehed", "toodete_arv", "artikli_ridade_arv", "kataloogi_kirjeldus"],
  groupDescriptions.map(([group, desc, pages]) => [group, pages, groupCount(group), articleCount(group), desc]),
  [180, 110, 95, 120, 520],
  "GrupidTable",
);

const controlRows = products.map((prod, index) => [
  index + 1,
  prod.page,
  prod.group,
  prod.product,
  prod.productId,
  prod.articleNumbers.length,
  "OK - tootekaart kataloogis",
]);
addTableSheet(
  workbook,
  "04_KONTROLL",
  ["nr", "kataloogi_leht", "tootegrupp", "toode", "product_id", "artikli_ridu", "kontrollistaatus"],
  controlRows,
  [55, 95, 165, 210, 135, 95, 190],
  "KontrollTable",
);

addTableSheet(
  workbook,
  "05_ALLIKAD",
  ["source_id", "allikas", "fail", "versioon", "lehed", "kasutus_failis", "märkus"],
  [
    [
      catalogSourceId,
      "Zenith Rubber Europe kataloog",
      catalogSourceFile,
      "2020 kataloog / PDF metadata: Zenith brochure 2020.pdf",
      "1-72",
      "Tarnijakataloog: tootegrupid, master-tooted, materjaliomadused, kasutusvaldkonnad, mõõdud ja artiklinumbrid.",
      "Kasutatakse Zenithist juurde tellimise ja tehnilise sobivuse kontrolliks. Iga toote ja artiklirea juures on allikas kujul fail + lk.",
    ],
    [
      "ZENITH_OLD_BASES",
      "Varasemad tööpõhjad samas kaustas",
      "Zenith_Materjalibaas_LOPLIK.xlsx; Zenith_Materjalibaas.sqlite; varasemad kummitoodete xlsx failid",
      "Kasutatud kontrolliks",
      "",
      "Abiallikad struktuuri ja varasemate ridade võrdlemiseks.",
      "Sisuline Zenith tooteloend on pandud kokku PDF-kataloogi põhjal.",
    ],
    [
      "ZENITH_PRICE_2025",
      "Zenith hinnakiri Plastokile",
      "Pricelist Zenith 2025 v2 (Plastok).xlsx",
      "2025 v2",
      "",
      "Zenithist juurde tellimise hinnad, allahindlused, hinnakirja mõõduvariandid ja artiklinumbrite hinna vaste.",
      "Nettohind on hinnakirja failis Plastokile arvestatud hind; sama artiklinumber võib korduda eri laiuse/pikkusega.",
    ],
    [
      "PLASTOK_DATABASE",
      "Plastok OÜ üldine materjalide andmebaas",
      "DATABASE.xlsm",
      "Lisatud kasutaja poolt",
      "",
      "Plastoki oma materjali-, kategooria-, hinna- ja laoseisu baas. Esmane kontroll enne Zenithist tellimist.",
      "Kui Plastoki oma materjal ei sobi või jääb puudu, kasutatakse Zenith kataloogi ja hinnakirja juurde tellimise otsuseks.",
    ],
  ],
  [135, 220, 360, 230, 80, 430, 330],
  "AllikadTable",
);

addTableSheet(
  workbook,
  "06_TOOTE_TUNNUSED",
  ["product_id", "tootegrupp", "toode", "tunnuse_tüüp", "tunnus", "andme_tase", "viidatud_allikas", "kontrolli_märkus"],
  tagRows,
  [115, 160, 190, 145, 240, 170, 220, 520],
  "TooteTunnusedTable",
);

addTableSheet(
  workbook,
  "07_MATERJALID",
  ["materjal_kvaliteet", "tootegrupid", "tooted", "tuletatud_kasutusfiltrid", "tuletatud_materjali_tunnused", "tuletatud_soovituslikud_kasutused", "kataloogi_lehed", "allikas"],
  materialRows,
  [140, 240, 420, 420, 420, 260, 110, 220],
  "MaterjalidTable",
);

addTableSheet(
  workbook,
  "08_VALIKUABI",
  [
    "kasutusolukord",
    "andme_tase",
    "soovituse_tase",
    "product_id",
    "toode",
    "tootegrupp",
    "ostjale",
    "insenerile",
    "tarbijale_muugiks",
    "kontrollida_enne_tellimist",
    "allikas",
    "soovituse_alus",
    "artiklid",
  ],
  selectionRows,
  [250, 170, 135, 115, 190, 170, 330, 420, 330, 360, 220, 420, 380],
  "ValikuabiTable",
);

addTableSheet(
  workbook,
  "09_ZENITH_HINNAD",
  [
    "article_nr",
    "name",
    "gross_price_eur",
    "nett_price_eur",
    "quantity",
    "unit",
    "thickness_mm",
    "width_mm",
    "length_mm",
    "group",
    "material",
    "min_temp_c",
    "max_temp_c",
    "colour",
    "hardness",
    "insertion",
    "insertion_type",
    "tensile_strength_mpa",
    "elongation_pct",
    "source_file",
  ],
  zenithPriceRows.map((row) => [
    row.article_nr,
    row.name,
    row.gross_price_eur,
    row.nett_price_eur,
    row.quantity,
    row.unit,
    row.thickness_mm,
    row.width_mm,
    row.length_mm,
    row.group,
    row.material,
    row.min_temp_c,
    row.max_temp_c,
    row.colour,
    row.hardness,
    row.insertion,
    row.insertion_type,
    row.tensile_strength_mpa,
    row.elongation_pct,
    row.source_file,
  ]),
  [145, 210, 105, 105, 85, 65, 95, 95, 95, 80, 100, 90, 90, 130, 90, 85, 120, 120, 105, 240],
  "ZenithHinnadTable",
);

addTableSheet(
  workbook,
  "10_ZENITH_HINNAKOOND",
  ["article_nr", "name", "hinnakirja_ridu", "unit", "gross_min_eur", "gross_max_eur", "nett_min_eur", "nett_max_eur", "widths_mm", "lengths_mm", "source_file"],
  zenithPriceSummaryRows.map((row) => [
    row.article_nr,
    row.name,
    row.rows,
    row.unit,
    row.gross_min_eur,
    row.gross_max_eur,
    row.nett_min_eur,
    row.nett_max_eur,
    row.widths_mm,
    row.lengths_mm,
    row.source_file,
  ]),
  [145, 210, 95, 80, 105, 105, 105, 105, 180, 180, 240],
  "ZenithHinnakoondTable",
);

addTableSheet(
  workbook,
  "11_ZENITH_ALLAHINDLUSED",
  ["article_or_group", "discount_pct", "article_group", "source_file"],
  zenithDiscountRows.map((row) => [row.article_or_group, row.discount_pct, row.article_group, row.source_file]),
  [220, 110, 120, 260],
  "ZenithAllahindlusedTable",
);

addTableSheet(
  workbook,
  "12_PLASTOK_DB_KOKKUVOTE",
  ["category", "group", "material", "ridu", "source_file"],
  plastokSummaryRows.map((row) => [row.category, row.group, row.material, row.rows, row.source_file]),
  [180, 120, 220, 90, 220],
  "PlastokDbKokkuvoteTable",
);

addTableSheet(
  workbook,
  "13_PLASTOK_DATABASE",
  [
    "article_nr",
    "name",
    "nett_price_eur",
    "gross_price_eur",
    "quantity",
    "unit",
    "category",
    "thickness",
    "width",
    "length",
    "group",
    "material",
    "min_temp_c",
    "max_temp_c",
    "colour",
    "hardness",
    "insertion",
    "insertion_type",
    "catalog",
    "tensile_strength_mpa",
    "elongation_pct",
    "source_file",
  ],
  plastokRows.map((row) => [
    row["Article nr."],
    row["Name"],
    row["Nett price, EUR"],
    row["Gross price, EUR"],
    row["Quantity"],
    row["Unit"],
    row["Category"],
    row["Thickness"],
    row["Width"],
    row["Length"],
    row["Group"],
    row["Material"],
    getAny(row, ["Min °C", "Min Â°C", "Min Ā°C"]),
    getAny(row, ["Max °C", "Max Â°C", "Max Ā°C"]),
    row["Colour"],
    row["Hardness"],
    row["Insertion"],
    row["Insertion type"],
    row["Catalog"],
    row["Tensile strength, Mpa"],
    row["Elongation, %"],
    "DATABASE.xlsm",
  ]),
  [130, 360, 105, 105, 90, 70, 150, 90, 90, 90, 90, 180, 90, 90, 120, 90, 85, 120, 100, 120, 105, 180],
  "PlastokDatabaseTable",
);

addTableSheet(
  workbook,
  "14_HINNA_VASTETA",
  ["rea_tüüp", "product_id", "toode", "tootegrupp", "article_nr", "toote_artiklid", "põhjus", "kataloogi_allikas"],
  missingPriceRows,
  [95, 120, 210, 170, 160, 420, 430, 220],
  "HinnaVastetaTable",
);

addTableSheet(
  workbook,
  "15_TÖÖPÕHIMÕTE",
  ["samm", "otsus", "kasutatav allikas", "mida kontrollida", "tulemus"],
  [
    [
      1,
      "Kontrolli esmalt Plastoki oma materjalibaasi.",
      "DATABASE.xlsm / 13_PLASTOK_DATABASE",
      "Materjal, mõõt, kogus/laoseis, hind, sertifikaat ja tehniline sobivus.",
      "Kui sobiv materjal on olemas, kasutatakse Plastoki oma materjali.",
    ],
    [
      2,
      "Kui Plastokil sobivat materjali, mõõtu või kogust ei ole, kontrolli Zenithit.",
      "Zenith catalogue 2020_0.pdf / 01_TOOTED / 02_ARTIKLID_MOODUD",
      "Sobiv materjal, kasutusvaldkond, tehnilised omadused, sertifikaadid ja artiklinumber.",
      "Zenith on juurde tellimise kanal, mitte oma materjalibaasi asendus.",
    ],
    [
      3,
      "Kui Zenith toode sobib, kontrolli hinda.",
      "Pricelist Zenith 2025 v2 (Plastok).xlsx / 09_ZENITH_HINNAD / 10_ZENITH_HINNAKOOND",
      "Nettohind, brutohind, ühik, mõõduvariant, artiklikoodi vaste ja allahindlus.",
      "Kui hind on olemas, saab toodet kasutada pakkumise sisendina.",
    ],
    [
      4,
      "Kui hinnavastet ei ole, ära eelda hinda.",
      "14_HINNA_VASTETA",
      "Kas kood on muutunud, toode on hinnakirjast puudu, toode on eritoode või vajab käsitsi päringut.",
      "Märgi rida ülevaatuseks või küsi Zenithilt/ostult hind üle.",
    ],
    [
      5,
      "Valikuabi on otsustugi, mitte automaatne lõplik tehniline kinnitus.",
      "08_VALIKUABI / 06_TOOTE_TUNNUSED / 07_MATERJALID",
      "Keskkond: õli, UV, temperatuur, toidukontakt, joogivesi, kulumine, elektriisolatsioon.",
      "Insener või müük kinnitab sobivuse kliendi tegeliku kasutusolukorra järgi.",
    ],
  ],
  [70, 300, 300, 420, 420],
  "ToopohimoteTable",
);

addTableSheet(
  workbook,
  "16_PLASTOK_MATERJALID",
  [
    "normaliseeritud_grupp",
    "andme_tase",
    "roll",
    "kasutusvaldkonnad",
    "tugevused",
    "piirangud",
    "ridu",
    "unikaalseid_artikleid",
    "kogusega_ridu",
    "bruto_min_eur",
    "bruto_max_eur",
    "netto_min_eur",
    "netto_max_eur",
    "ühikud",
    "kategooriad",
    "grupid",
    "materjali_koodid",
    "näidisartiklid",
    "millal_zenith",
    "allikas",
  ],
  plastokMaterialRows,
  [190, 190, 280, 380, 420, 420, 70, 115, 95, 95, 95, 95, 95, 120, 240, 160, 360, 340, 420, 180],
  "PlastokMaterjalidTable",
);

addTableSheet(
  workbook,
  "17_PLASTOK_ENNE_ZENITH",
  [
    "kasutusolukord",
    "andme_tase",
    "plastoki_esimesed_grupid",
    "plastoki_oma_baasi_kokkuvõte",
    "oma_materjal_piisab_kui",
    "millal_tellida_zenithist",
    "zenith_sobivad_tooted",
    "kontrollida",
    "risk",
    "allikad",
  ],
  plastokFirstDecisionRows,
  [260, 170, 290, 500, 420, 430, 340, 380, 360, 330],
  "PlastokEnneZenithTable",
);

addTableSheet(
  workbook,
  "18_ANDMETÄPSUS_AUDIT",
  [
    "product_id",
    "tootegrupp",
    "toode",
    "kataloogi_leht",
    "pdf_tootenimi_leitud",
    "pdf_tootenimi_ocr_variant",
    "pdf_product_id_leitud",
    "pdf_product_id_ocr_variant",
    "kataloogi_artikleid",
    "pdf_artikleid_leitud",
    "pdf_artikleid_kontrollida",
    "hinnakirjas_artikleid_leitud",
    "hinnakirjas_kontrollida",
    "pdf_kontrollida_artiklid",
    "hinnakirja_kontrollida_artiklid",
    "staatus",
    "kontrolli_märkus",
    "allikas",
  ],
  productAuditRows,
  [115, 160, 190, 90, 125, 145, 125, 145, 110, 115, 125, 150, 135, 420, 420, 240, 380, 220],
  "AndmetapsusAuditTable",
);

addTableSheet(
  workbook,
  "19_ARTIKLI_AUDIT",
  [
    "product_id",
    "tootegrupp",
    "toode",
    "kataloogi_leht",
    "article_nr",
    "pdf_tootenimi_leitud",
    "pdf_tootenimi_ocr_variant",
    "pdf_artiklikood_leitud",
    "pdf_artiklikood_ocr_variant",
    "hinnakirja_vaste_leitud",
    "hinnakirja_vasteid",
    "staatus",
    "kontrolli_märkus",
    "allikas",
  ],
  articleAuditRows,
  [115, 160, 190, 90, 160, 130, 155, 140, 160, 145, 120, 300, 380, 220],
  "ArtikliAuditTable",
);

addTableSheet(
  workbook,
  "20_ALLIKA_KONTROLL",
  ["kontrollitav_allikas", "fail", "kontroll", "leitud", "oodatud_või_võrdlus", "staatus"],
  sourceControlRows(),
  [210, 300, 360, 120, 180, 180],
  "AllikaKontrollTable",
);

addTableSheet(
  workbook,
  "21_KÄSITSI_KONTROLL",
  ["rea_tüüp", "product_id", "toode", "article_nr", "allikas", "põhjus", "kontrolli_juhis", "staatus"],
  manualControlRows,
  [120, 120, 210, 170, 220, 420, 420, 300],
  "KasitsiKontrollTable",
);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const overview = await workbook.inspect({
  kind: "table",
  range: "00_KOKKUVOTE!A1:F51",
  include: "values",
  tableMaxRows: 51,
  tableMaxCols: 6,
});
console.log(overview.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const renderTargets = [
  { sheetName: "00_KOKKUVOTE", range: "A1:F51" },
  { sheetName: "01_TOOTED", range: "A1:AK40" },
  { sheetName: "02_ARTIKLID_MOODUD", range: "A1:Y40" },
  { sheetName: "03_GRUPID", range: "A1:E15" },
  { sheetName: "04_KONTROLL", range: "A1:G40" },
  { sheetName: "05_ALLIKAD", range: "A1:G8" },
  { sheetName: "06_TOOTE_TUNNUSED", range: "A1:H40" },
  { sheetName: "07_MATERJALID", range: "A1:H25" },
  { sheetName: "08_VALIKUABI", range: "A1:M35" },
  { sheetName: "09_ZENITH_HINNAD", range: "A1:T40" },
  { sheetName: "10_ZENITH_HINNAKOOND", range: "A1:K40" },
  { sheetName: "11_ZENITH_ALLAHINDLUSED", range: "A1:D40" },
  { sheetName: "12_PLASTOK_DB_KOKKUVOTE", range: "A1:E40" },
  { sheetName: "13_PLASTOK_DATABASE", range: "A1:V40" },
  { sheetName: "14_HINNA_VASTETA", range: "A1:H40" },
  { sheetName: "15_TÖÖPÕHIMÕTE", range: "A1:E8" },
  { sheetName: "16_PLASTOK_MATERJALID", range: "A1:T40" },
  { sheetName: "17_PLASTOK_ENNE_ZENITH", range: "A1:J15" },
  { sheetName: "18_ANDMETÄPSUS_AUDIT", range: "A1:R40" },
  { sheetName: "19_ARTIKLI_AUDIT", range: "A1:N40" },
  { sheetName: "20_ALLIKA_KONTROLL", range: "A1:F12" },
  { sheetName: "21_KÄSITSI_KONTROLL", range: "A1:H40" },
];
for (const target of renderTargets) {
  const preview = await workbook.render({
    sheetName: target.sheetName,
    ...(target.range ? { range: target.range } : { autoCrop: "all" }),
    scale: 1,
    format: "png",
  });
  await fs.writeFile(path.join(outputDir, `${target.sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(
  JSON.stringify(
    {
      outputPath,
      productCount: products.length,
      variantCount: variantRows.length,
      groupCount: groupDescriptions.length,
      zenithPriceRows: zenithPriceRows.length,
      plastokRows: plastokRows.length,
      plastokMaterialGroups: plastokMaterialRows.length,
      plastokFirstDecisionRows: plastokFirstDecisionRows.length,
      pdfProductNameFound: `${pdfProductNameFoundCount}/${products.length}`,
      pdfArticleFound: `${pdfArticleFoundCount}/${variantRows.length}`,
      priceArticleFound: `${priceArticleFoundCount}/${variantRows.length}`,
      priceArticleMissing: priceArticleMissingCount,
    },
    null,
    2,
  ),
);
