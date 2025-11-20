const vengeanceSalvationSplashSynergy = +config.vengeanceSalvationSplashSynergy;
const vengeanceResistSplashSynergy = +config.vengeanceResistSplashSynergy;
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

        if (vengeanceExplosionRadius) {
            // Turn Vengeance into a "melee + missiles" skill
            row.auralencalc = 60;
            row.cltdofunc = 29;

            // All missiles point to the custom AoE missile
            row.cltmissile  = "vengeancemissile";
            row.srvmissile  = "vengeancemissile";
        }
    }
});

D2RMM.writeTsv(skillsFilename, skills);

// ===== skilldesc.txt =====
const skilldescFilename = 'global\\excel\\skilldesc.txt';
const skilldesc = D2RMM.readTsv(skilldescFilename);

skilldesc.rows.forEach((row) => {
    if (row.skilldesc === "vengeance") {
        // Add Sacrifice Attack Rating synergy text
        row.dsc3line6  = 76;
        row.dsc3texta6 = "AttRateplev";
        row.dsc3textb6 = "skillname96";
        row.dsc3calca6 = vengeanceSacrificeArSynergy;
        row.dsc3calcb6 = "";

        if (vengeanceExplosionRadius) {
            row.dsc2line1 = 36;
            row.dsc2texta1 = 'StrSkillRadiusSingular';
            row.dsc2textb1 = '';
            row.dsc2calca1 = `${vengeanceExplosionRadius}`;
            row.dsc2calcb1 = '';

            row.descmissile1 = "vengeancemissile_fire";
            row.descmissile2 = "vengeancemissile_cold";
            row.descmissile3 = "vengeancemissile_ltng";

            row.descline3 = 75;
            row.desctexta3 = "VengeanceElemDmgAoE";
            row.desctextb3 = "";
            row.desccalca3 = "m1en + m2en + m3en";
            row.desccalcb3 = "m1ex + m2ex + m3ex";
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
        ResultFlags: 5,
        HitFlags: 2,
        HitShift: 8,
        NumDirections: 1,

        EType: elem,
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
        MaxELev5: 8,

        // Cold gets chill length; others don't
        ELen: (elem === "cold" ? 75 : ""),

        // Synergies with Resist Cold/Fire/Lightning
        EDmgSymPerCalc: elem === 'cold' ? `(skill('Resist Cold'.blvl))*${vengeanceResistSplashSynergy} + (skill('Salvation'.blvl))*${vengeanceSalvationSplashSynergy}` : elem === 'fire'
            ? `(skill('Resist Fire'.blvl))*${vengeanceResistSplashSynergy} + (skill('Salvation'.blvl))*${vengeanceSalvationSplashSynergy}`
            : `(skill('Resist Lightning'.blvl))*${vengeanceResistSplashSynergy} + (skill('Salvation'.blvl))*${vengeanceSalvationSplashSynergy}`,
        "*eol": 0
    });
}

D2RMM.writeTsv(missilesFilename, missiles);

// ===== skills.json =====
const skillsJsonFileName = 'local\\lng\\strings\\skills.json';
const skillsJson = D2RMM.readJson(skillsJsonFileName);
let skillIds = skillsJson.map((row) => row.id);
let skillId = 92040;
while (skillIds.includes(skillId)) {
    skillId++;
}
console.log(`Adding skill ${skillId} to skills.json`);

skillsJson.push(
    {
        "id": skillId,
        "Key": "VengeanceElemDmgAoE",
        "enUS": "Elemental Explosion Damage: %d-%d",
        "zhTW": "元素爆炸傷害: %d-%d",
        "deDE": "Elementarer Explosionsschaden: %d-%d",
        "esES": "Daño de explosión elemental: %d-%d",
        "frFR": "Dégâts d'explosion élémentaire : %d-%d",
        "itIT": "Danni da esplosione elementale: %d-%d",
        "koKR": "원소 폭발 피해: %d-%d",
        "plPL": "Obrażenia od wybuchu żywiołów: %d-%d",
        "esMX": "Daño de explosión elemental: %d-%d",
        "jaJP": "元素爆発ダメージ: %d-%d",
        "ptBR": "Dano de Explosão Elemental: %d-%d",
        "ruRU": "Урон от стихийного взрыва: %d-%d",
        "zhCN": "元素爆炸伤害: %d-%d"
    }
);

D2RMM.writeJson(skillsJsonFileName, skillsJson);