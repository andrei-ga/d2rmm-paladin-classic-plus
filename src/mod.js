const vengeanceSalvationDmgSynergy = +config.vengeanceSalvationDmgSynergy;
const vengeanceSalvationSplashSynergy = +config.vengeanceSalvationSplashSynergy;
const vengeanceSacrificeArSynergy = +config.vengeanceSacrificeArSynergy;
const vengeanceExplosionRadius = +config.vengeanceExplosionRadius;

// ===== skills.txt =====
const skillsFilename = 'global\\excel\\skills.txt';
const skills = D2RMM.readTsv(skillsFilename);

skills.rows.forEach((row) => {
    if (row.skill === "Vengeance") {
        // Read the *original* AR progression so we don't break it
        const baseToHit = row.ToHit;      // base % AR
        const baseLevToHit = row.LevToHit; // % AR per Vengeance level

        // Use ToHitCalc to override the built-in formula:
        // original AR + % per Sacrifice hard point
        row.ToHitCalc = `${baseToHit} + (lvl-1)*${baseLevToHit} + skill('Sacrifice'.blvl)*${vengeanceSacrificeArSynergy}`;

        // Make Salvation the only synergy for elemental % (simple version)
        row.Param7 = vengeanceSalvationDmgSynergy;
        row.calc1 = "ln12+skill('Salvation'.blvl)*par7";
        row.calc2 = "ln12+skill('Salvation'.blvl)*par7";
        row.calc3 = "ln12+skill('Salvation'.blvl)*par7";

        if (vengeanceExplosionRadius) {
            // Turn Vengeance into a "melee + missiles" skill
            row.auralencalc = 60;
            row.cltdofunc = 29;

            // All missiles point to the custom AoE missile
            row.cltmissile  = "vengeancemissile";
            row.cltmissilea = "vengeancemissile";
            row.cltmissileb = "vengeancemissile";
            row.cltmissilec = "vengeancemissile";

            row.srvmissile  = "vengeancemissile";
            row.srvmissilea = "vengeancemissile";
            row.srvmissileb = "vengeancemissile";
            row.srvmissilec = "vengeancemissile";
        }
    }
});

D2RMM.writeTsv(skillsFilename, skills);

// ===== skilldesc.txt =====
const skilldescFilename = 'global\\excel\\skilldesc.txt';
const skilldesc = D2RMM.readTsv(skilldescFilename);

skilldesc.rows.forEach((row) => {
    if (row.skilldesc === "vengeance") {
        // Adjust skill text to show correct calculation
        row.desccalca2 = "ln12+skill('Salvation'.blvl)*par7";
        row.desccalca4 = "ln12+skill('Salvation'.blvl)*par7";
        row.desccalca5 = "ln12+skill('Salvation'.blvl)*par7";

        // Remove Resist Passive Synergies
        row.dsc3line3 = "";
        row.dsc3texta3 = "";
        row.dsc3textb3 = "";
        row.dsc3calca3 = "";

        row.dsc3line4 = "";
        row.dsc3texta4 = "";
        row.dsc3textb4 = "";
        row.dsc3calca4 = "";

        // Add Salvation first row
        row.dsc3line2 = 76;
        row.dsc3texta2 = "ElemDampLev";
        row.dsc3textb2 = "skillname125";
        row.dsc3calca2 = "par7";

        // Remove previous Salvation row
        row.dsc3line5 = "";
        row.dsc3texta5 = "";
        row.dsc3textb5 = "";
        row.dsc3calca5 = "";

        // Add missiles for calculation references
        row.descmissile1 = "vengeancemissile_fire";
        row.descmissile2 = "vengeancemissile_cold";
        row.descmissile3 = "vengeancemissile_ltng";

        // Add AoE dmg numbers in place of Cold Length
        row.descline3 = 75;
        row.desctexta3 = "VengeanceElemDmgAoE";
        row.desctextb3 = "";
        row.desccalca3 = "m1en";
        row.desccalcb3 = "m1ex";

        // Add Salvation Synergy text for AoE Damage
        row.dsc3line3 = 76;
        row.dsc3texta3 = "VengeanceElemDmgAoESynergy";
        row.dsc3textb3 = "skillname125";
        row.dsc3calca3 = vengeanceSalvationSplashSynergy;
        row.dsc3calcb3 = "";

        // Add Sacrifice Attack Rating synergy text
        row.dsc3line4  = 76;
        row.dsc3texta4 = "VengeanceARSacrificeSynergy";
        row.dsc3textb4 = "skillname96";
        row.dsc3calca4 = vengeanceSacrificeArSynergy;
        row.dsc3calcb4 = "";

        if (vengeanceExplosionRadius) {
            row.dsc2line1 = 36;
            row.dsc2texta1 = 'StrSkillRadiusSingular';
            row.dsc2textb1 = '';
            row.dsc2calca1 = `${vengeanceExplosionRadius}`;
            row.dsc2calcb1 = '';
        }
    }
});

D2RMM.writeTsv(skilldescFilename, skilldesc);

// ===== missiles.txt =====
const missilesFilename = 'global\\excel\\missiles.txt';
const missiles = D2RMM.readTsv(missilesFilename);
let missileID = Math.max(...missiles.rows.map((row) => row['*ID']));

// Base AoE container missile
missiles.rows.push({
    "*ID": ++missileID,
    Missile: "vengeancemissile",
    pCltDoFunc: 1,
    pSrvDoFunc: 37,
    pSrvHitFunc: 1,
    sHitPar1: 5,
    Range: 1,
    CelFile: "null",
    animrate: 1024,
    AnimLen: 16,
    AnimSpeed: 16,
    CollideType: 3,
    CollideKill: 1,
    Size: 1,
    AlwaysExplode: 1,
    GetHit: 1,
    Trans: 1,
    MissileSkill: 1,
    ResultFlags: 5,
    HitFlags: 2,
    HitShift: 8,
    EType: 'frze',
    NumDirections: 1,
    SubMissile1: "vengeancemissile_cold",
    SubMissile2: "vengeancemissile_fire",
    SubMissile3: "vengeancemissile_ltng",
    "*eol": 0
});
for (let elem of ["cold", "fire", "ltng"]) {
    missiles.rows.push({
        "*ID": ++missileID,
        Missile: `vengeancemissile_${elem}`,
        pCltDoFunc: 1,
        pSrvDoFunc: 1,
        pSrvHitFunc: 1,
        sHitPar1: vengeanceExplosionRadius, // Hit radius
        Range: 1,
        CelFile: "null",
        animrate: 1024,
        AnimLen: 16,
        AnimSpeed: 16,
        CollideType: 3,
        CollideKill: 1,
        Size: 1,
        AlwaysExplode: 1,
        GetHit: 1,
        Trans: 1,
        MissileSkill: 1,
        ResultFlags: 5,
        HitFlags: 2,
        HitShift: 8,
        NumDirections: 1,
        EType: elem,

        // Base ele damage and level scaling
        EMin: 4,
        MinELev1: 3,
        MinELev2: 4,
        MinELev3: 5,
        MinELev4: 6,
        MinELev5: 7,

        EMax: 6,
        MaxELev1: 3,
        MaxELev2: 4,
        MaxELev3: 5,
        MaxELev4: 6,
        MaxELev5: 9,

        // Cold gets chill length; others don't
        ELen: (elem === "cold" ? 75 : ""),

        // Synergy: Salvation increases splash damage
        EDmgSymPerCalc: `(skill('Salvation'.blvl))*${vengeanceSalvationSplashSynergy}`,
        "*eol": 0
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
        "zhTW": "火焰/閃電/冰冷範圍傷害：%d-%d",
        "deDE": "Feuer-/Blitz-/Kälte-Flächenschaden: %d-%d",
        "esES": "Daño de salpicadura de Fuego/Rayo/Frío: %d-%d",
        "frFR": "Dégâts de zone Feu/Foudre/Froid : %d-%d",
        "itIT": "Danni ad area Fuoco/Fulmine/Freddo: %d-%d",
        "koKR": "화염/번개/냉기 범위 피해: %d-%d",
        "plPL": "Obrażenia obszarowe od Ognia/Błyskawic/Zimna: %d-%d",
        "esMX": "Daño de área de Fuego/Rayo/Frío: %d-%d",
        "jaJP": "火炎/稲妻/冷気の範囲ダメージ: %d-%d",
        "ptBR": "Dano em área de Fogo/Raio/Frio: %d-%d",
        "ruRU": "Урон по области огнём/молнией/холодом: %d–%d",
        "zhCN": "火焰/闪电/冰冷范围伤害：%d-%d"
    },
    {
        "id": ++skillId,
        "Key": "VengeanceElemDmgAoESynergy",
        "enUS": "%s: %+d%% Fire/Light/Cold Splash Damage per Level",
        "zhTW": "%s：每級火焰/閃電/冰冷範圍傷害 %+d%%",
        "deDE": "%s: %+d%% Feuer-/Blitz-/Kälte-Flächenschaden pro Stufe",
        "esES": "%s: %+d%% de daño de salpicadura de Fuego/Rayo/Frío por nivel",
        "frFR": "%s : %+d%% de dégâts de zone Feu/Foudre/Froid par niveau",
        "itIT": "%s: %+d%% danni ad area Fuoco/Fulmine/Freddo per livello",
        "koKR": "%s: 레벨당 화염/번개/냉기 범위 피해 %+d%%",
        "plPL": "%s: %+d%% obrażeń obszarowych od Ognia/Błyskawic/Zimna za poziom",
        "esMX": "%s: %+d%% de daño de área de Fuego/Rayo/Frío por nivel",
        "jaJP": "%s: レベルにつき火炎/稲妻/冷気の範囲ダメージ %+d%%",
        "ptBR": "%s: %+d%% de dano em área de Fogo/Raio/Frio por nível",
        "ruRU": "%s: %+d%% урона по области огнём/молнией/холодом за уровень",
        "zhCN": "%s：每级火焰/闪电/冰冷范围伤害 %+d%%"
    },
    {
        "id": ++skillId,
        "Key": "VengeanceARSacrificeSynergy",
        "enUS": "%s: +%d%% Attack Rating per Level",
        "zhTW": "%s：每級 +%d%% 攻擊命中率",
        "deDE": "%s: +%d%% Angriffswert pro Stufe",
        "esES": "%s: +%d%% al índice de ataque por nivel",
        "frFR": "%s : +%d%% à l'indice d'attaque par niveau",
        "itIT": "%s: +%d%% all'indice d'attacco per livello",
        "koKR": "%s: 레벨당 명중률 +%d%%",
        "plPL": "%s: +%d%% do współczynnika trafienia za poziom",
        "esMX": "%s: +%d%% a la puntería de ataque por nivel",
        "jaJP": "%s: レベルにつき攻撃命中率 +%d%%",
        "ptBR": "%s: +%d%% na Taxa de Ataque por nível",
        "ruRU": "%s: +%d%% к шансу попадания за уровень",
        "zhCN": "%s：每级 +%d%% 攻击命中率"
    }
);

D2RMM.writeJson(skillsJsonFileName, skillsJson);