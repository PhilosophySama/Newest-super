const fs = require('fs');
const path = require('path');

function parseStageAutomation(content) {
    const triggers = [];
    const mappingsMatch = content.match(/STAGE_MAPPINGS:\s*{([^}]+)}/s);
    
    if (mappingsMatch) {
        const mappingsStr = mappingsMatch[1];
        const lines = mappingsStr.split('\n').filter(l => l.includes(':'));
        
        lines.forEach(line => {
            const match = line.match(/(\w+):\s*\[([^\]]+)\]/);
            if (match) {
                const triggerKey = match[1];
                const wordsMatch = match[2].match(/'([^']+)'/g);
                const words = wordsMatch ? wordsMatch.map(w => w.replace(/'/g, '')) : [];
                
                triggers.push({
                    id: triggerKey,
                    type: 'stage_move',
                    primaryWord: words[0] || triggerKey,
                    allWords: words,
                    file: 'Stage Automation.gs',
                    lineNumber: findLineNumber(content, triggerKey + ':')
                });
            }
        });
    }
    
    return triggers;
}

function parseDraftCreator(content) {
    const configs = [
        { key: 'TARGET_STAGE', defaultValue: 'test' },
        { key: 'LIZ_STAGE', defaultValue: 'Liz' },
        { key: 'CUSTOMER_STAGE', defaultValue: 'Email customer' },
        { key: 'HANDOFF_STAGE', defaultValue: 'Cust Handoff' },
        { key: 'ROUGH_QUOTE_STAGE', defaultValue: 'Rough quote' },
        { key: 'COI_STAGE', defaultValue: 'COI Req' }
    ];
    
    return configs.map(config => {
        const match = content.match(new RegExp(`${config.key}:\\s*'([^']+)'`));
        const word = match ? match[1] : config.defaultValue;
        
        return {
            id: config.key,
            type: 'draft_creator',
            primaryWord: word,
            allWords: [word],
            file: 'Draft Creator.gs',
            lineNumber: match ? findLineNumber(content, match[0]) : 0
        };
    });
}

function parseAutoTriggers() {
    return [
        { id: 'col_f', column: 'F', name: 'Display Name', sheets: ['Leads', 'F/U', 'Awarded'] },
        { id: 'col_i', column: 'I', name: 'Email', sheets: ['Leads', 'F/U', 'Awarded'] },
        { id: 'col_j', column: 'J', name: 'Address', sheets: ['Leads', 'F/U', 'Awarded'] },
        { id: 'col_m', column: 'M', name: 'Job Description', sheets: ['Leads'] },
        { id: 'col_p', column: 'P', name: 'QB URL', sheets: ['All'] },
        { id: 'col_q', column: 'Q', name: 'Earth Link', sheets: ['Leads', 'F/U', 'Awarded'] },
        { id: 'col_a', column: 'A', name: 'Split CSV', sheets: ['Leads'] }
    ].map(t => ({ ...t, type: 'auto_column', file: 'Stage Automation.gs' }));
}

function findLineNumber(content, searchStr) {
    const lines = content.split('\n');
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes(searchStr)) return i + 1;
    }
    return 0;
}

function parseAllScripts() {
    const rootDir = __dirname; // Read from root, not scripts folder
    const allTriggers = {
        stageMoves: [],
        draftCreator: [],
        autoColumns: [],
        metadata: {
            lastUpdated: new Date().toISOString(),
            totalTriggers: 0,
            version: '1.0',
            repoUrl: 'https://github.com/PhilosophySama/Current-Appscript-files'
        }
    };
    
    try {
        const stageContent = fs.readFileSync(path.join(rootDir, 'Stage Automation.gs'), 'utf8');
        allTriggers.stageMoves = parseStageAutomation(stageContent);
        console.log(`✅ Parsed Stage Automation.gs: ${allTriggers.stageMoves.length} triggers`);
    } catch (err) {
        console.error('⚠️ Could not parse Stage Automation.gs:', err.message);
    }
    
    try {
        const draftContent = fs.readFileSync(path.join(rootDir, 'Draft Creator.gs'), 'utf8');
        allTriggers.draftCreator = parseDraftCreator(draftContent);
        console.log(`✅ Parsed Draft Creator.gs: ${allTriggers.draftCreator.length} triggers`);
    } catch (err) {
        console.error('⚠️ Could not parse Draft Creator.gs:', err.message);
    }
    
    allTriggers.autoColumns = parseAutoTriggers();
    console.log(`✅ Added ${allTriggers.autoColumns.length} auto-column triggers`);
    
    allTriggers.metadata.totalTriggers = 
        allTriggers.stageMoves.length + 
        allTriggers.draftCreator.length + 
        allTriggers.autoColumns.length;
    
    const outputPath = path.join(rootDir, 'docs', 'triggers.json');
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, JSON.stringify(allTriggers, null, 2));
    
    console.log(`\n🎉 Successfully parsed ${allTriggers.metadata.totalTriggers} total triggers`);
    console.log(`📝 Output saved to: ${outputPath}`);
    return allTriggers;
}

// Run parser
if (require.main === module) {
    parseAllScripts();
}

module.exports = { parseAllScripts };
