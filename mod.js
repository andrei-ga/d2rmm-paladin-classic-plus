const vengeanceSalvationDmgSynergy = +config.vengeanceSalvationDmgSynergy;
const vengeanceSalvationSplashSynergy = +config.vengeanceSalvationSplashSynergy;
const vengeanceSacrificeArSynergy = +config.vengeanceSacrificeArSynergy;

// ===== skills.txt =====
const skillsFilename = 'global\\excel\\skills.txt';
const skills = D2RMM.readTsv(skillsFilename);

skills.rows.forEach((row) => {
    if (row["skill"] === "Vengeance") {
        // Slightly cheaper / adjusted Vengeance
        row["manashift"] = 5;

        // Read the *original* AR progression so we don't break it
        const baseToHit = row['ToHit'] || '0';      // base % AR
        const baseLevToHit = row['LevToHit'] || '0'; // % AR per Vengeance level

        // Use ToHitCalc to override the built-in formula:
        // original AR + % per Sacrifice hard point
        row['ToHitCalc'] = `${baseToHit} + (lvl-1)*${baseLevToHit} + skill('Sacrifice'.blvl)*${vengeanceSacrificeArSynergy}`;

        // Make Salvation the only synergy for elemental % (simple version)
        row["Param7"] = vengeanceSalvationDmgSynergy;
        row["calc1"] = "ln12+skill('Salvation'.blvl)*par7";
        row["calc2"] = "ln12+skill('Salvation'.blvl)*par7";
        row["calc3"] = "ln12+skill('Salvation'.blvl)*par7";

        // Turn Vengeance into a "melee + missiles" skill
        row["auralencalc"] = 60;
        row["cltdofunc"] = 29;

        // All missiles point to the custom AoE missile
        row["cltmissile"]  = "vengeancemissile";
        row["cltmissilea"] = "vengeancemissile";
        row["cltmissileb"] = "vengeancemissile";
        row["cltmissilec"] = "vengeancemissile";

        row["srvmissile"]  = "vengeancemissile";
        row["srvmissilea"] = "vengeancemissile";
        row["srvmissileb"] = "vengeancemissile";
        row["srvmissilec"] = "vengeancemissile";
    }
});

D2RMM.writeTsv(skillsFilename, skills);

// ===== skilldesc.txt =====
const skilldescFilename = 'global\\excel\\skilldesc.txt';
const skilldesc = D2RMM.readTsv(skilldescFilename);

skilldesc.rows.forEach((row) => {
    if (row["skilldesc"] === "vengeance") {
        // Adjust skill text to show correct calculation
        row["desccalca2"] = "ln12+skill('Salvation'.blvl)*par7";
        row["desccalca4"] = "ln12+skill('Salvation'.blvl)*par7";
        row["desccalca5"] = "ln12+skill('Salvation'.blvl)*par7";

        // Remove Resist Passive Synergies
        row["dsc3line3"] = "";
        row["dsc3texta3"] = "";
        row["dsc3textb3"] = "";
        row["dsc3calca3"] = "";

        row["dsc3line4"] = "";
        row["dsc3texta4"] = "";
        row["dsc3textb4"] = "";
        row["dsc3calca4"] = "";

        // Add Salvation first row
        row["dsc3line2"] = 76;
        row["dsc3texta2"] = "ElemDampLev";
        row["dsc3textb2"] = "skillname125";
        row["dsc3calca2"] = "par7";

        // Remove previous Salvation row
        row["dsc3line5"] = "";
        row["dsc3texta5"] = "";
        row["dsc3textb5"] = "";
        row["dsc3calca5"] = "";

        // Add missiles for calculation references
        row["descmissile1"] = "vengeancemissile_fire";
        row["descmissile2"] = "vengeancemissile_cold";
        row["descmissile3"] = "vengeancemissile_ltng";

        // Add AoE dmg numbers in place of Cold Length
        row["descline3"] = 75;
        row["desctexta3"] = "VengeanceElemDmgAoE";
        row["desctextb3"] = "";
        row["desccalca3"] = "m1en";
        row["desccalcb3"] = "m1ex";

        // Add Salvation Synergy text for AoE Damage
        row["dsc3line3"] = 76;
        row["dsc3texta3"] = "VengeanceElemDmgAoESynergy";
        row["dsc3textb3"] = "skillname125";
        row["dsc3calca3"] = vengeanceSalvationSplashSynergy;
        row["dsc3calcb3"] = "";

        // Add Sacrifice Attack Rating synergy text
        row["dsc3line4"]  = 76;
        row["dsc3texta4"] = "VengeanceARSacrificeSynergy";
        row["dsc3textb4"] = "skillname96";
        row["dsc3calca4"] = vengeanceSacrificeArSynergy;
        row["dsc3calcb4"] = "";
    }
});

D2RMM.writeTsv(skilldescFilename, skilldesc);

// ===== missiles.txt =====
const missilesFilename = 'global\\excel\\missiles.txt';
const missiles = D2RMM.readTsv(missilesFilename);
let missileID = Math.max(...missiles.rows.map((row) => row['*ID']));

// Base AoE container missile
const freezingArrowExp3 = missiles.rows.find((row) => row["Missile"] === "freezingarrowexp3");
missiles.rows.push({
    ...freezingArrowExp3,
    "*ID": ++missileID,
    "Missile": "vengeancemissile",
    "MissileSkill": "",
    "Skill": "",
    "pCltDoFunc": 1,
    "pSrvDoFunc": 37,
    "pSrvHitFunc": 1,
    "SubMissile1": "vengeancemissile_cold",
    "SubMissile2": "vengeancemissile_fire",
    "SubMissile3": "vengeancemissile_ltng"
});
for (let elem of ["cold", "fire", "ltng"]) {
    missiles.rows.push({
        ...freezingArrowExp3,
        "*ID": ++missileID,
        "Missile": `vengeancemissile_${elem}`,
        "MissileSkill": "",
        "Skill": "",
        "pCltDoFunc": 1,
        "pSrvDoFunc": 1,
        "pSrvHitFunc": 1,
        "sHitPar1": 10, // Hit radius
        "EType": elem,

        // Base ele damage and level scaling
        "EMin": 4,
        "MinELev1": 3,
        "MinELev2": 4,
        "MinELev3": 5,
        "MinELev4": 6,
        "MinELev5": 7,

        "EMax": 6,
        "MaxELev1": 3,
        "MaxELev2": 4,
        "MaxELev3": 5,
        "MaxELev4": 6,
        "MaxELev5": 9,

        "SubMissile1": "",
        "SubMissile2": "",
        "SubMissile3": "",

        // Cold gets chill length; others don't
        "ELen": (elem === "cold" ? 75 : ""),

        // Synergy: Salvation increases splash damage
        "EDmgSymPerCalc": `(skill('Salvation'.blvl))*${vengeanceSalvationSplashSynergy}`
    });
}

D2RMM.writeTsv(missilesFilename, missiles);

// ===== skills.json =====
const skillsJsonFileName = 'local\\lng\\strings\\skills.json';
const skillsJson = D2RMM.readJson(skillsJsonFileName);
let skillId = Math.max(...skillsJson.map((row) => row['id']));

skillsJson.push(
    {
        "id": ++skillId,
        "Key": "VengeanceElemDmgAoE",
        "enUS": "Fire/Light/Cold Splash Damage: %d-%d",
        "zhTW": "Fire/Light/Cold Splash Damage: %d-%d",
        "deDE": "Fire/Light/Cold Splash Damage: %d-%d",
        "esES": "Fire/Light/Cold Splash Damage: %d-%d",
        "frFR": "Fire/Light/Cold Splash Damage: %d-%d",
        "itIT": "Fire/Light/Cold Splash Damage: %d-%d",
        "koKR": "Fire/Light/Cold Splash Damage: %d-%d",
        "plPL": "Fire/Light/Cold Splash Damage: %d-%d",
        "esMX": "Fire/Light/Cold Splash Damage: %d-%d",
        "jaJP": "Fire/Light/Cold Splash Damage: %d-%d",
        "ptBR": "Fire/Light/Cold Splash Damage: %d-%d",
        "ruRU": "Fire/Light/Cold Splash Damage: %d-%d",
        "zhCN": "Fire/Light/Cold Splash Damage: %d-%d"
    }
);

skillsJson.push(
    {
        "id": ++skillId,
        "Key": "VengeanceElemDmgAoESynergy",
        "enUS": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "zhTW": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "deDE": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "esES": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "frFR": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "itIT": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "koKR": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "plPL": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "esMX": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "jaJP": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "ptBR": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "ruRU": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "zhCN": "%s: %+d%% Fire/Light/Cold Splash Damage per Level"
    }
);

skillsJson.push(
    {
        "id": ++skillId,
        "Key": "VengeanceARSacrificeSynergy",
        "enUS": "%s: +%d%% Attack Rating per Level",
        "zhTW": "%s: +%d%% Attack Rating per Level",
        "deDE": "%s: +%d%% Attack Rating per Level",
        "esES": "%s: +%d%% Attack Rating per Level",
        "frFR": "%s: +%d%% Attack Rating per Level",
        "itIT": "%s: +%d%% Attack Rating per Level",
        "koKR": "%s: +%d%% Attack Rating per Level",
        "plPL": "%s: +%d%% Attack Rating per Level",
        "esMX": "%s: +%d%% Attack Rating per Level",
        "jaJP": "%s: +%d%% Attack Rating per Level",
        "ptBR": "%s: +%d%% Attack Rating per Level",
        "ruRU": "%s: +%d%% Attack Rating per Level",
        "zhCN": "%s: +%d%% Attack Rating per Level"
    }
);

D2RMM.writeJson(skillsJsonFileName, skillsJson);