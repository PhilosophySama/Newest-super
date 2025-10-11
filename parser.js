const fs = require('fs');
const path = require('path');

// ============================================
// COMPREHENSIVE .GS FILE PARSER
// Extracts: triggers, functions, dependencies, purposes
// ============================================

function getAllGsFiles(dir, fileList = []) {
    const files = fs.readdirSync(dir);
    
    files.forEach(file => {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);
        
        if (stat.isDirectory()) {
            // Recursively scan subdirectories
            getAllGsFiles(filePath, fileList);
        } else if (file.endsWith('.gs')) {
            fileList.push(filePath);
        }
    });
    
    return fileList;
}

function extractFilePurpose(content) {
    // Extract from header comments
    const lines = content.split('\n').slice(0, 20);
    const comments = [];
    
    for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith('/**') || trimmed.startsWith('/*')) continue;
        if (trimmed.startsWith('*') && !trimmed.startsWith('*/')) {
            const comment = trimmed.replace(/^\*\s*/, '').trim();
            if (comment && !comment.includes('====') && comment.length > 5) {
                comments.push(comment);
            }
        }
        if (trimmed.startsWith('*/')) break;
    }
    
    return comments.slice(0, 3).join(' ') || 'No description available';
}

function extractVersion(content) {
    const versionMatch = content.match(/[Vv]ersion[:#\s]+([^\n]+)/);
    return versionMatch ? versionMatch[1].trim() : null;
}

function extractFunctions(content) {
    const functions = [];
    const functionPattern = /function\s+(\w+)\s*\([^)]*\)\s*{/g;
    let match;
    
    while ((match = functionPattern.exec(content)) !== null) {
        const funcName = match[1];
        const lineNum = content.substring(0, match.index).split('\n').length;
        
        // Get preceding comment if exists
        const lines = content.substring(0, match.index).split('\n');
        let description = '';
        for (let i = lines.length - 1; i >= Math.max(0, lines.length - 5); i--) {
            const line = lines[i].trim();
            if (line.startsWith('//') || line.startsWith('*')) {
                description = line.replace(/^[\/\*\s]+/, '') + ' ' + description;
            } else if (line === '') {
                continue;
            } else {
                break;
            }
        }
        
        functions.push({
            name: funcName,
            line: lineNum,
            description: description.trim() || null,
            isPublic: !funcName.endsWith('_'),
            isHelper: funcName.includes('_')
        });
    }
    
    return functions;
}

function extractTriggers(content, filename) {
    const triggers = [];
    
    // Stage mappings (Stage Automation style)
    if (filename.includes('Stage Automation')) {
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
                        type: 'stage_mapping',
                        key: triggerKey,
                        words: words,
                        primary: words[0] || triggerKey
                    });
                }
            });
        }
    }
    
    // Draft stages (Draft Creator style)
    if (filename.includes('Draft Creator')) {
        const stagePatterns = [
            { key: 'TARGET_STAGE', default: 'test' },
            { key: 'LIZ_STAGE', default: 'Liz' },
            { key: 'CUSTOMER_STAGE', default: 'Email customer' },
            { key: 'HANDOFF_STAGE', default: 'Cust Handoff' },
            { key: 'ROUGH_QUOTE_STAGE', default: 'Rough quote' },
            { key: 'COI_STAGE', default: 'COI Req' }
        ];
        
        stagePatterns.forEach(pattern => {
            const match = content.match(new RegExp(`${pattern.key}:\\s*'([^']+)'`));
            const word = match ? match[1] : pattern.default;
            
            triggers.push({
                type: 'draft_stage',
                key: pattern.key,
                words: [word],
                primary: word
            });
        });
    }
    
    // Column-based triggers
    const colTriggers = [
        { pattern: /handleDisplayNameChange_/, column: 'F', name: 'Display Name' },
        { pattern: /handleEmailChange_/, column: 'I', name: 'Email' },
        { pattern: /handleAddressChange_/, column: 'J', name: 'Address' },
        { pattern: /updateJobDescription_/, column: 'M', name: 'Job Description' },
        { pattern: /handleQbUrlChange_/, column: 'P', name: 'QB URL' },
        { pattern: /createEarthLink_/, column: 'Q', name: 'Earth Link' }
    ];
    
    colTriggers.forEach(ct => {
        if (ct.pattern.test(content)) {
            triggers.push({
                type: 'column_trigger',
                column: ct.column,
                name: ct.name,
                automated: true
            });
        }
    });
    
    return triggers;
}

function extractDependencies(content) {
    const deps = [];
    
    // Function calls to other files
    const callPatterns = [
        /(\w+)_\(/g,  // Helper functions with underscore
        /MOVE_CONFIG/g,
        /DRAFTS_V2/g,
        /LEANTO_RUBY_EXPORT/g,
        /MILEAGE_CONFIG/g
    ];
    
    callPatterns.forEach(pattern => {
        const matches = content.match(pattern);
        if (matches) {
            matches.forEach(m => {
                const cleaned = m.replace(/[()]/g, '');
                if (cleaned && !deps.includes(cleaned)) {
                    deps.push(cleaned);
                }
            });
        }
    });
    
    // Config object references
    const configMatch = content.match(/const\s+(\w+_CONFIG)\s*=/);
    if (configMatch) {
        deps.push('Defines: ' + configMatch[1]);
    }
    
    return [...new Set(deps)];
}

function parseGsFile(filePath) {
    const content = fs.readFileSync(filePath, 'utf8');
    const relativePath = path.relative(__dirname, filePath);
    const filename = path.basename(filePath);
    
    return {
        path: relativePath,
        filename: filename,
        purpose: extractFilePurpose(content),
        version: extractVersion(content),
        functions: extractFunctions(content),
        triggers: extractTriggers(content, filename),
        dependencies: extractDependencies(content),
        lines: content.split('\n').length,
        size: content.length
    };
}

function categorizeFiles(files) {
    const categories = {
        core: [],
        automation: [],
        integration: [],
        utilities: [],
        other: []
    };
    
    files.forEach(file => {
        const name = file.filename.toLowerCase();
        
        if (name.includes('stage automation') || name.includes('draft creator')) {
            categories.core.push(file);
        } else if (name.includes('mileage') || name.includes('lean-to') || name.includes('ruby')) {
            categories.automation.push(file);
        } else if (name.includes('quickbooks') || name.includes('qbo')) {
            categories.integration.push(file);
        } else if (name.includes('menu') || name.includes('test')) {
            categories.utilities.push(file);
        } else {
            categories.other.push(file);
        }
    });
    
    return categories;
}

function generateStats(files) {
    let totalFunctions = 0;
    let totalTriggers = 0;
    let totalLines = 0;
    let publicFunctions = 0;
    let helperFunctions = 0;
    
    files.forEach(file => {
        totalFunctions += file.functions.length;
        totalTriggers += file.triggers.length;
        totalLines += file.lines;
        
        file.functions.forEach(func => {
            if (func.isPublic) publicFunctions++;
            if (func.isHelper) helperFunctions++;
        });
    });
    
    return {
        totalFiles: files.length,
        totalFunctions,
        totalTriggers,
        totalLines,
        publicFunctions,
        helperFunctions,
        avgFunctionsPerFile: Math.round(totalFunctions / files.length),
        avgLinesPerFile: Math.round(totalLines / files.length)
    };
}

// ============================================
// MAIN PARSER
// ============================================

function parseAllScripts() {
    console.log('🔍 Scanning for .gs files...\n');
    
    const rootDir = __dirname;
    const gsFiles = getAllGsFiles(rootDir);
    
    console.log(`Found ${gsFiles.length} .gs files:`);
    gsFiles.forEach(f => console.log(`  - ${path.relative(rootDir, f)}`));
    console.log('');
    
    // Parse each file
    const parsedFiles = gsFiles.map(parseGsFile);
    
    // Categorize
    const categories = categorizeFiles(parsedFiles);
    
    // Generate stats
    const stats = generateStats(parsedFiles);
    
    // Compile output
    const output = {
        metadata: {
            lastUpdated: new Date().toISOString(),
            repoUrl: 'https://github.com/PhilosophySama/Newest-super',
            version: '2.0-complete',
            scanType: 'comprehensive'
        },
        stats,
        files: parsedFiles,
        categories,
        // Legacy format for backward compatibility
        stageMoves: parsedFiles.flatMap(f => 
            f.triggers.filter(t => t.type === 'stage_mapping')
        ),
        draftCreator: parsedFiles.flatMap(f => 
            f.triggers.filter(t => t.type === 'draft_stage')
        ),
        autoColumns: parsedFiles.flatMap(f => 
            f.triggers.filter(t => t.type === 'column_trigger')
        )
    };
    
    // Legacy total count
    output.metadata.totalTriggers = 
        output.stageMoves.length + 
        output.draftCreator.length + 
        output.autoColumns.length;
    
    // Save output
    const outputPath = path.join(rootDir, 'docs', 'triggers.json');
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, JSON.stringify(output, null, 2));
    
    // Print summary
    console.log('\n✅ PARSING COMPLETE!\n');
    console.log('📊 Statistics:');
    console.log(`   Files scanned:        ${stats.totalFiles}`);
    console.log(`   Total functions:      ${stats.totalFunctions}`);
    console.log(`   Public functions:     ${stats.publicFunctions}`);
    console.log(`   Helper functions:     ${stats.helperFunctions}`);
    console.log(`   Total triggers:       ${stats.totalTriggers}`);
    console.log(`   Total lines of code:  ${stats.totalLines.toLocaleString()}`);
    console.log(`\n📝 Output saved to: ${outputPath}`);
    console.log(`\n🌐 View at: https://philosophysama.github.io/Newest-super/\n`);
    
    return output;
}

// Run if called directly
if (require.main === module) {
    parseAllScripts();
}

module.exports = { parseAllScripts };
