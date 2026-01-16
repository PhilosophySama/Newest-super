/**
 * LEAN-TO & A-FRAME RUBY (Exact) — Drive File Export
 * Version: 1/16 9am EST by Claude Sonnet 4.5
 *
 * Triggers on edits to Leads!T:AD (cols 20–30) when AA contains:
 *   - "Lean-to" or "Sloped L" → generates Lean-to Ruby
 *   - "A-frame" or "A frame" → generates A-frame Ruby
 *
 * Generates plain-text .rb files in Drive with parameters:
 *   - AWNING_MATERIAL (AB)
 *   - LENGTH (T)
 *   - PROJECTION (U)
 *   - HEIGHT (X)
 *   - FRONT_BAR_HEIGHT (V)
 *   - HAS_WINGS (Y > 0 → true, else false)
 *   - HAS_POSTS (AD == "Yes" → true; checkbox TRUE → true; otherwise false)
 *
 * Column S shows a hyperlink "Ruby (.rb)". If Drive write fails, falls back to writing into S.
 *
 * Helper names prefixed r_ to avoid collisions.
 */

const AWNING_RUBY_EXPORT = {
  FOLDER_ID: '',                        // optional fixed folder ID (leave blank to auto-create in My Drive)
  FOLDER_NAME: 'SketchUp Ruby Exports', // used when FOLDER_ID is blank; created once and cached
  LINK_TEXT: 'Ruby (.rb)'
};

function handleEditAwningRuby_(e) {
  if (!e || !e.source || !e.range) return;

  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Leads') return;

  // Only single-cell data-row edits
  const r = e.range;
  const row = r.getRow();
  const col = r.getColumn();
  if (row === 1 || r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;

  // Only react to T–AD (20–30)
  if (col < 20 || col > 30) return;

  const COLS = {
    RUBY_OUT: 19,      // S
    LENGTH: 20,        // T
    WIDTH: 21,         // U
    FRONT_BAR: 22,     // V
    WING_HEIGHT: 24,   // X
    NUM_WINGS: 25,     // Y  (0 => false, else true)
    TYPE: 27,          // AA
    FABRIC: 28,        // AB
    POSTS: 30,         // AD  ("Yes" or checkbox TRUE → true; else false)
    DISPLAY: 6         // F - optional for nicer filenames
  };

  const type = String(sheet.getRange(row, COLS.TYPE).getDisplayValue() || '').trim().toLowerCase();
  
  // Determine awning type
  let awningType = null;
  if (type && (type.includes('lean') || type === 'sloped l')) {
    awningType = 'LEAN_TO';
  } else if (type && (type.includes('a-frame') || type.includes('a frame'))) {
    awningType = 'A_FRAME';
  } else {
    sheet.getRange(row, COLS.RUBY_OUT).clearContent();
    return;
  }

  // Read parameters
  const length    = Number(sheet.getRange(row, COLS.LENGTH).getValue())      || 50;
  const width     = Number(sheet.getRange(row, COLS.WIDTH).getValue())       || 11;
  const height    = Number(sheet.getRange(row, COLS.WING_HEIGHT).getValue()) || 5;
  const frontBar  = Number(sheet.getRange(row, COLS.FRONT_BAR).getValue())   || 1;
  const wingsNum  = Number(sheet.getRange(row, COLS.NUM_WINGS).getValue())   || 0;

  const fabricRaw = sheet.getRange(row, COLS.FABRIC).getValue();   // could be text/formula
  const postsVal  = sheet.getRange(row, COLS.POSTS).getValue();    // could be text/checkbox boolean
  const dispName  = String(sheet.getRange(row, COLS.DISPLAY).getDisplayValue() || '').trim();

  const hasWings   = wingsNum > 0;
  const fabricType = (typeof fabricRaw === 'string' && /vinyl/i.test(fabricRaw)) ? 'Vinyl' : 'Sunbrella';

  // Robust posts flag:
  //  - TRUE if checkbox is checked (boolean true)
  //  - TRUE if text equals "yes" (case-insensitive, trimmed)
  //  - FALSE for blank, false, anything else
  let hasPosts = false;
  if (typeof postsVal === 'boolean') {
    hasPosts = postsVal === true;
  } else if (postsVal != null) {
    hasPosts = /^\s*yes\s*$/i.test(String(postsVal));
  }

  // Build the Ruby script based on awning type
  let rubyExact;
  let typeName;
  
  if (awningType === 'LEAN_TO') {
    rubyExact = r_buildRubyFromLeanToTemplate_({
      fabric:     fabricType,
      length:     length,
      projection: width,
      height:     height,
      frontBar:   frontBar,
      hasWings:   hasWings,
      hasPosts:   hasPosts
    });
    typeName = 'Lean-to';
  } else if (awningType === 'A_FRAME') {
    rubyExact = r_buildRubyFromAFrameTemplate_({
      fabric:     fabricType,
      length:     length,
      projection: width,
      height:     height,
      frontBar:   frontBar,
      hasWings:   hasWings,
      hasPosts:   hasPosts
    });
    typeName = 'A-Frame';
  }

  try {
    const file = r_writeRubyToDrive_(rubyExact, row, dispName, typeName);
    const url  = file.getUrl();

    const rich = SpreadsheetApp.newRichTextValue()
      .setText(AWNING_RUBY_EXPORT.LINK_TEXT)
      .setLinkUrl(url)
      .build();
    sheet.getRange(row, COLS.RUBY_OUT).setRichTextValue(rich);

    SpreadsheetApp.getActive().toast('Ruby export ready for row ' + row, 'Success', 3);
  } catch (err) {
    sheet.getRange(row, COLS.RUBY_OUT).setValue(rubyExact); // fallback
    SpreadsheetApp.getActive().toast('Drive write failed, wrote to S instead: ' + err.message, 'Warning', 5);
  }
}

/**
 * Replace ONLY the parameter literals in the Lean-to Ruby template.
 * Every other character (including spacing and comments) is preserved verbatim.
 */
function r_buildRubyFromLeanToTemplate_(p) {
  let txt = r_getLeanToRubyTemplate_();

  // AWNING_MATERIAL
  txt = txt.replace(
    /^(AWNING_MATERIAL\s*=\s*)"(?:[^"]*)"(\s*#.*)$/m,
    `$1"${p.fabric}"$2`
  );

  // LENGTH
  txt = txt.replace(
    /^(LENGTH\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.length}$2`
  );

  // PROJECTION
  txt = txt.replace(
    /^(PROJECTION\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.projection}$2`
  );

  // HEIGHT
  txt = txt.replace(
    /^(HEIGHT\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.height}$2`
  );

  // FRONT_BAR_HEIGHT
  txt = txt.replace(
    /^(FRONT_BAR_HEIGHT\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.frontBar}$2`
  );

  // HAS_WINGS (true/false)
  txt = txt.replace(
    /^(HAS_WINGS\s*=\s*)(true|false)(\s*#.*)$/m,
    `$1${p.hasWings}$3`
  );

  // HAS_POSTS (true/false)
  txt = txt.replace(
    /^(HAS_POSTS\s*=\s*)(true|false)(\s*#.*)$/m,
    `$1${p.hasPosts}$3`
  );

  return txt;
}

/**
 * Replace ONLY the parameter literals in the A-Frame Ruby template.
 * Every other character (including spacing and comments) is preserved verbatim.
 */
function r_buildRubyFromAFrameTemplate_(p) {
  let txt = r_getAFrameRubyTemplate_();

  // AWNING_MATERIAL
  txt = txt.replace(
    /^(AWNING_MATERIAL\s*=\s*)"(?:[^"]*)"(\s*#.*)$/m,
    `$1"${p.fabric}"$2`
  );

  // LENGTH
  txt = txt.replace(
    /^(LENGTH\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.length}$2`
  );

  // PROJECTION (divide by 2 for A-frame since spreadsheet has total width)
  txt = txt.replace(
    /^(PROJECTION\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.projection / 2}$2`
  );

  // PEAK_HEIGHT (calculated from height + frontBar)
  const peakHeight = p.height + p.frontBar;
  txt = txt.replace(
    /^(PEAK_HEIGHT\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${peakHeight}$2`
  );

  // FRONT_BAR_HEIGHT
  txt = txt.replace(
    /^(FRONT_BAR_HEIGHT\s*=\s*)[0-9.]+(\s*#.*)$/m,
    `$1${p.frontBar}$2`
  );

  // NUM_WINGS
  txt = txt.replace(
    /^(NUM_WINGS\s*=\s*)[0-9]+(\s*#.*)$/m,
    `$1${p.hasWings ? 2 : 0}$2`
  );

  // HAS_POSTS (true/false)
  txt = txt.replace(
    /^(HAS_POSTS\s*=\s*)(true|false)(\s*#.*)$/m,
    `$1${p.hasPosts}$3`
  );

  return txt;
}

/**
 * LEAN-TO RUBY TEMPLATE — DO NOT CHANGE ANYTHING HERE.
 * We substitute only the parameter values via r_buildRubyFromLeanToTemplate_().
 */
function r_getLeanToRubyTemplate_() {
  return String.raw`# Ver: ChatGPT - 09/13 - 02:25 AM - Lean-to Awnings (Material Spacing + Swapped Wings + 2x2 Posts)

# === AWNING CONFIGURATION ===
AWNING_TYPE          = "Lean-to"    # Label for awning type
AWNING_MATERIAL      = "Sunbrella"  # Options: "Sunbrella" or "Vinyl"
LENGTH               = 50           # Length (along x) in feet
PROJECTION           = 11           # Width / Projection (along y) in feet  
HEIGHT               = 5            # Wall side height (z) in feet
FRONT_BAR_HEIGHT     = 1            # Front bar height (z) in feet
HAS_WINGS            = true         # true = left/right wings, false = no wings
HAS_DIAGONAL_BRACING = true         # true = diagonal bracing inside wings
HAS_POSTS            = false        # true = add vertical posts, false = no posts
COLUMN_HEIGHT        = 7            # column height in feet (default 7')
POST_SIZE            = 2.0          # post size in inches (square cross-section)

# === SCRIPT BEGINS - DO NOT MODIFY BELOW ===
model = Sketchup.active_model
entities = model.active_entities
model.start_operation("Create #{AWNING_TYPE} Awning", true)

# Convert feet → inches
length           = LENGTH * 12
projection       = PROJECTION * 12
height           = HEIGHT * 12
front_bar_height = FRONT_BAR_HEIGHT * 12
col_height       = COLUMN_HEIGHT * 12

# Determine rafter spacing (in inches) based on material
spacing_ft = if AWNING_MATERIAL.downcase == "sunbrella"
               3.6
             else
               5.0
             end
spacing_in = spacing_ft * 12

# Compute number of spans/rafters
num_spans     = (length / spacing_in).ceil
num_supports  = num_spans + 1
spacing       = length.to_f / num_spans

# Create group for everything
awning_group    = entities.add_group
group_entities  = awning_group.entities

# === MAIN SLOPE PLANE ===
back_left   = [0, 0, height]
back_right  = [length, 0, height]
front_left  = [0, projection, front_bar_height]
front_right = [length, projection, front_bar_height]

awning_face = group_entities.add_face(back_left, back_right, front_right, front_left)

# === RAFTERS ===
(0...num_supports).each do |i|
  x_pos = [i * spacing, length].min
  group_entities.add_line([x_pos, 0, height], [x_pos, projection, front_bar_height])
end

# === FRONT BAR VERTICALS ===
num_verticals = (num_spans * 2) + 1
(0...num_verticals).each do |i|
  x_pos = [i * (spacing / 2), length].min
  group_entities.add_line([x_pos, projection, front_bar_height], [x_pos, projection, 0])
end

# === FRONT BAR RECTANGLE ===
# Top beam
group_entities.add_line([0, projection, front_bar_height], [length, projection, front_bar_height])
# Bottom beam
group_entities.add_line([length, projection, 0], [0, projection, 0])

# Front bar face (vertical plane)
group_entities.add_face(
  [0, projection, front_bar_height],
  [length, projection, front_bar_height],
  [length, projection, 0],
  [0, projection, 0]
)


# === FRONT POSTS (2"x2" COLUMNS, in separate sub-group) ===
if HAS_POSTS
  # Create a sub-group for posts
  posts_group = group_entities.add_group
  posts_entities = posts_group.entities
  
  col_height = COLUMN_HEIGHT * 12
  post_spacing_sections = (length / (15.0 * 12)).ceil  # number of ~15' sections
  num_posts = post_spacing_sections + 1
  spacing_x = length.to_f / post_spacing_sections

  post_size = POST_SIZE # in inches

  (0..post_spacing_sections).each do |i|
    x_pos = [i * spacing_x, length].min

    # Ensure post stays inside X and Y boundaries
    x1 = [x_pos, length - post_size].min
    x2 = [x1 + post_size, length].min
    y1 = projection - post_size
    y2 = projection

    pt1 = Geom::Point3d.new(x1, y1, 0)
    pt2 = Geom::Point3d.new(x2, y1, 0)
    pt3 = Geom::Point3d.new(x2, y2, 0)
    pt4 = Geom::Point3d.new(x1, y2, 0)

    face = posts_entities.add_face(pt1, pt2, pt3, pt4)
    face.reverse! if face.normal.z < 0

    # Extrude downwards by column height
    face.pushpull(-col_height)
  end
  
  posts_group.name = "Support Posts"
end

# === WALL ATTACHMENT LINE ===
group_entities.add_line([0, 0, height], [length, 0, height])

# === WINGS ===
# Note: Wings named from user perspective facing the building
# LEFT WING is at x=length, RIGHT WING is at x=0
if HAS_WINGS
  # LEFT WING (x = length)
  left_mid_y = projection / 2.0
  left_mid_z = height - ((height - front_bar_height) * (left_mid_y / projection))
  group_entities.add_line([length, left_mid_y, left_mid_z], [length, left_mid_y, 0])
  group_entities.add_line([length, 0, 0], [length, projection, 0])  # base line

  if HAS_DIAGONAL_BRACING
    group_entities.add_line([length, 0, 0], [length, left_mid_y, left_mid_z])
  end

  group_entities.add_face(
    [length, 0, height],
    [length, projection, front_bar_height],
    [length, projection, 0],
    [length, 0, 0]
  )

  # RIGHT WING (x = 0)
  right_mid_y = projection / 2.0
  right_mid_z = height - ((height - front_bar_height) * (right_mid_y / projection))
  group_entities.add_line([0, right_mid_y, right_mid_z], [0, right_mid_y, 0])
  group_entities.add_line([0, 0, 0], [0, projection, 0])  # base line

  if HAS_DIAGONAL_BRACING
    group_entities.add_line([0, 0, 0], [0, right_mid_y, right_mid_z])
  end

  group_entities.add_face(
    [0, 0, height],
    [0, projection, front_bar_height],
    [0, projection, 0],
    [0, 0, 0]
  )
end

# === DIMENSIONS (inside group) ===
# Length x
group_entities.add_dimension_linear(
  Geom::Point3d.new(0, 0, height),
  Geom::Point3d.new(length, 0, height),
  Geom::Vector3d.new(0,0,24)
)
# Projection y
group_entities.add_dimension_linear(
  Geom::Point3d.new(length, 0, 0),
  Geom::Point3d.new(length, projection, 0),
  Geom::Vector3d.new(0,0,-24)
)
# Height z
group_entities.add_dimension_linear(
  Geom::Point3d.new(length, 0, 0),
  Geom::Point3d.new(length, 0, height),
  Geom::Vector3d.new(24,0,0)
)
# Front bar height
group_entities.add_dimension_linear(
  Geom::Point3d.new(0, projection, 0),
  Geom::Point3d.new(0, projection, front_bar_height),
  Geom::Vector3d.new(0,24,0)
)

# === NAME GROUP + COMMIT ===
awning_group.name = "#{AWNING_TYPE} (#{AWNING_MATERIAL}) #{LENGTH}x#{PROJECTION}x#{HEIGHT}-FB#{FRONT_BAR_HEIGHT}-POST#{COLUMN_HEIGHT}"
model.commit_operation`;
}

/**
 * A-FRAME RUBY TEMPLATE — DO NOT CHANGE ANYTHING HERE.
 * We substitute only the parameter values via r_buildRubyFromAFrameTemplate_().
 */
function r_getAFrameRubyTemplate_() {
  return String.raw`# A-Frame Awning Generator
# Ver: Claude Opus 4.1 - 01/13 - 11:45 PM EST

# === AWNING CONFIGURATION ===
AWNING_TYPE          = "A-Frame"    # Label for awning type
AWNING_MATERIAL      = "Sunbrella"  # Options: "Sunbrella" or "Vinyl"
LENGTH               = 20           # Length (along x) in feet
PROJECTION           = 5            # Width / Projection per side (along y) in feet  
PEAK_HEIGHT          = 4            # Peak height (z) in feet
FRONT_BAR_HEIGHT     = 1            # Front bar height (z) in feet
NUM_WINGS            = 2            # Wings on both ends (0 or 2)
HAS_POSTS            = false        # true = add vertical posts, false = no posts
COLUMN_HEIGHT        = 7            # column height in feet (default 7')
POST_SIZE            = 2.0          # post size in inches (square cross-section)

# === SCRIPT BEGINS - DO NOT MODIFY BELOW ===
model = Sketchup.active_model
entities = model.active_entities
model.start_operation("Create #{AWNING_TYPE} Awning", true)

# Convert feet → inches
length           = LENGTH * 12
projection       = PROJECTION * 12
peak_height      = PEAK_HEIGHT * 12
front_bar_height = FRONT_BAR_HEIGHT * 12
col_height       = COLUMN_HEIGHT * 12

# Determine truss spacing
num_trusses   = 4 + 1  # Number of trusses
truss_spacing = length.to_f / 4

# Create main group
aframe_group = entities.add_group
group_entities = aframe_group.entities

# === A-FRAME GEOMETRY ===
# Peak runs along the x-axis at center
peak_left = [0, 0, peak_height]
peak_right = [length, 0, peak_height]
ground_left_front = [0, -projection, front_bar_height]
ground_left_back = [length, -projection, front_bar_height]

# Right slope (positive y direction)
ground_right_front = [0, projection, front_bar_height]
ground_right_back = [length, projection, front_bar_height]

# Create left slope face (normal should point outward/left = negative y)
left_face = group_entities.add_face(
  peak_left, ground_left_front, ground_left_back, peak_right
)
left_face.reverse! if left_face.normal.y > 0

# Create right slope face (normal should point outward/right = positive y)
right_face = group_entities.add_face(
  peak_left, peak_right, ground_right_back, ground_right_front
)
right_face.reverse! if right_face.normal.y < 0

# === TRUSSES (RAFTERS) ===
(0...num_trusses).each do |i|
  x_pos = [i * truss_spacing, length].min
  # Left rafter
  group_entities.add_line([x_pos, 0, peak_height], [x_pos, -projection, front_bar_height])
  # Right rafter
  group_entities.add_line([x_pos, 0, peak_height], [x_pos, projection, front_bar_height])
end

# Peak line
group_entities.add_line(peak_left, peak_right)

# === FRONT BAR VERTICALS (LEFT SIDE) ===
num_verticals = (4 * 2) + 1
(0...num_verticals).each do |i|
  x_pos = [i * (truss_spacing / 2), length].min
  group_entities.add_line([x_pos, -projection, front_bar_height], [x_pos, -projection, 0])
end

# === FRONT BAR VERTICALS (RIGHT SIDE) ===
(0...num_verticals).each do |i|
  x_pos = [i * (truss_spacing / 2), length].min
  group_entities.add_line([x_pos, projection, front_bar_height], [x_pos, projection, 0])
end

# === FRONT BAR RECTANGLES ===
# Left side front bar (normal should point outward/left = negative y)
# Top beam
group_entities.add_line([0, -projection, front_bar_height], [length, -projection, front_bar_height])
# Bottom beam
group_entities.add_line([0, -projection, 0], [length, -projection, 0])
# Face
left_bar_face = group_entities.add_face(
  [0, -projection, front_bar_height],
  [0, -projection, 0],
  [length, -projection, 0],
  [length, -projection, front_bar_height]
)
left_bar_face.reverse! if left_bar_face.normal.y > 0

# Right side front bar (normal should point outward/right = positive y)
# Top beam
group_entities.add_line([0, projection, front_bar_height], [length, projection, front_bar_height])
# Bottom beam
group_entities.add_line([0, projection, 0], [length, projection, 0])
# Face
right_bar_face = group_entities.add_face(
  [0, projection, front_bar_height],
  [length, projection, front_bar_height],
  [length, projection, 0],
  [0, projection, 0]
)
right_bar_face.reverse! if right_bar_face.normal.y < 0

# === FRONT POSTS (2"x2" COLUMNS, in separate sub-group) ===
if HAS_POSTS
  # Create sub-groups for posts
  posts_group_left = group_entities.add_group
  posts_entities_left = posts_group_left.entities
  
  posts_group_right = group_entities.add_group
  posts_entities_right = posts_group_right.entities
  
  post_spacing_sections = (length / (15.0 * 12)).ceil  # number of ~15' sections
  spacing_x = length.to_f / post_spacing_sections
  post_size = POST_SIZE # in inches

  # Left side posts
  (0..post_spacing_sections).each do |i|
    x_pos = [i * spacing_x, length].min

    # Ensure post stays inside X and Y boundaries
    x1 = [x_pos, length - post_size].min
    x2 = [x1 + post_size, length].min
    y1 = -projection
    y2 = -projection + post_size

    pt1 = Geom::Point3d.new(x1, y1, 0)
    pt2 = Geom::Point3d.new(x2, y1, 0)
    pt3 = Geom::Point3d.new(x2, y2, 0)
    pt4 = Geom::Point3d.new(x1, y2, 0)

    face = posts_entities_left.add_face(pt1, pt2, pt3, pt4)
    face.reverse! if face.normal.z < 0

    # Extrude downwards by column height
    face.pushpull(-col_height)
  end
  
  # Right side posts
  (0..post_spacing_sections).each do |i|
    x_pos = [i * spacing_x, length].min

    # Ensure post stays inside X and Y boundaries
    x1 = [x_pos, length - post_size].min
    x2 = [x1 + post_size, length].min
    y1 = projection - post_size
    y2 = projection

    pt1 = Geom::Point3d.new(x1, y1, 0)
    pt2 = Geom::Point3d.new(x2, y1, 0)
    pt3 = Geom::Point3d.new(x2, y2, 0)
    pt4 = Geom::Point3d.new(x1, y2, 0)

    face = posts_entities_right.add_face(pt1, pt2, pt3, pt4)
    face.reverse! if face.normal.z < 0

    # Extrude downwards by column height
    face.pushpull(-col_height)
  end
  
  posts_group_left.name = "Support Posts (Left)"
  posts_group_right.name = "Support Posts (Right)"
end

# === WINGS ===
if NUM_WINGS > 0
  # Left end wing (x = 0) - normal should point in negative x direction
  left_mid_y_neg = -projection / 2.0
  left_mid_z_neg = peak_height - ((peak_height - front_bar_height) * 0.5)
  group_entities.add_line([0, left_mid_y_neg, left_mid_z_neg], [0, left_mid_y_neg, 0])
  group_entities.add_line([0, -projection, 0], [0, 0, 0])
  
  left_mid_y_pos = projection / 2.0
  left_mid_z_pos = peak_height - ((peak_height - front_bar_height) * 0.5)
  group_entities.add_line([0, left_mid_y_pos, left_mid_z_pos], [0, left_mid_y_pos, 0])
  group_entities.add_line([0, 0, 0], [0, projection, 0])
  
  # Left wing faces
  left_wing_neg = group_entities.add_face([0, 0, peak_height], [0, 0, 0], [0, -projection, 0], [0, -projection, front_bar_height])
  left_wing_neg.reverse! if left_wing_neg.normal.x > 0
  
  left_wing_pos = group_entities.add_face([0, 0, peak_height], [0, projection, front_bar_height], [0, projection, 0], [0, 0, 0])
  left_wing_pos.reverse! if left_wing_pos.normal.x > 0
  
  # Right end wing (x = LENGTH) - normal should point in positive x direction
  right_mid_y_neg = -projection / 2.0
  right_mid_z_neg = peak_height - ((peak_height - front_bar_height) * 0.5)
  group_entities.add_line([length, right_mid_y_neg, right_mid_z_neg], [length, right_mid_y_neg, 0])
  group_entities.add_line([length, -projection, 0], [length, 0, 0])
  
  right_mid_y_pos = projection / 2.0
  right_mid_z_pos = peak_height - ((peak_height - front_bar_height) * 0.5)
  group_entities.add_line([length, right_mid_y_pos, right_mid_z_pos], [length, right_mid_y_pos, 0])
  group_entities.add_line([length, 0, 0], [length, projection, 0])
  
  # Right wing faces
  right_wing_neg = group_entities.add_face([length, 0, peak_height], [length, -projection, front_bar_height], [length, -projection, 0], [length, 0, 0])
  right_wing_neg.reverse! if right_wing_neg.normal.x < 0
  
  right_wing_pos = group_entities.add_face([length, 0, peak_height], [length, 0, 0], [length, projection, 0], [length, projection, front_bar_height])
  right_wing_pos.reverse! if right_wing_pos.normal.x < 0
end

# Delete bottom plane if it exists
group_entities.grep(Sketchup::Face).each do |face|
  if face.normal.z.abs > 0.99 && face.vertices.all? { |v| v.position.z.abs < 0.01 }
    face.erase!
  end
end

# === DIMENSIONS ===
group_entities.add_dimension_linear(
  Geom::Point3d.new(0, 0, peak_height),
  Geom::Point3d.new(length, 0, peak_height),
  Geom::Vector3d.new(0, 0, 24)
)

# Projection dimension (left side)
group_entities.add_dimension_linear(
  Geom::Point3d.new(length, 0, 0),
  Geom::Point3d.new(length, -projection, 0),
  Geom::Vector3d.new(24, 0, 0)
)

# Height dimension
group_entities.add_dimension_linear(
  Geom::Point3d.new(0, 0, 0),
  Geom::Point3d.new(0, 0, peak_height),
  Geom::Vector3d.new(0, -projection - 24, 0)
)

# Front bar height
group_entities.add_dimension_linear(
  Geom::Point3d.new(0, -projection, 0),
  Geom::Point3d.new(0, -projection, front_bar_height),
  Geom::Vector3d.new(0, -12, 0)
)

# === NAME GROUP + COMMIT ===
aframe_group.name = "#{AWNING_TYPE} (#{AWNING_MATERIAL}) #{LENGTH}x#{PROJECTION}x#{PEAK_HEIGHT}-FB#{FRONT_BAR_HEIGHT}-POST#{COLUMN_HEIGHT}"
model.commit_operation`;
}

/** Write the Ruby to Drive as a .rb file and return the File. */
function r_writeRubyToDrive_(rubyScript, row, dispName, typeName) {
  const folder = r_getRubyExportFolder_();
  const safeName = (dispName || ('Row ' + row))
    .replace(/[\\/:*?"<>|#\[\]\r\n]+/g, ' ')
    .trim()
    .substring(0, 80);

  const filename = `${typeName} Ruby - ${safeName || ('Row ' + row)}.rb`;

  // Remove any old file(s) with same name (keeps folder tidy)
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const blob = Utilities.newBlob(rubyScript, 'text/plain', filename);
  return folder.createFile(blob);
}

/** Get/create the export folder once and cache its ID. */
function r_getRubyExportFolder_() {
  if (AWNING_RUBY_EXPORT.FOLDER_ID) {
    return DriveApp.getFolderById(AWNING_RUBY_EXPORT.FOLDER_ID);
  }
  const prop = PropertiesService.getScriptProperties();
  const cached = prop.getProperty('AWNING_RUBY_EXPORT_FOLDER_ID');
  if (cached) {
    try { return DriveApp.getFolderById(cached); } catch (_) { /* recreate below */ }
  }
  const folder = DriveApp.createFolder(AWNING_RUBY_EXPORT.FOLDER_NAME);
  prop.setProperty('AWNING_RUBY_EXPORT_FOLDER_ID', folder.getId());
  return folder;
}

/** Trigger installer (use from menu: Setup → Install On-Edit Trigger (Awning Ruby)) */
function installTriggerAwningRuby_() {
  const ssId = SpreadsheetApp.getActive().getId();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleEditAwningRuby_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEditAwningRuby_').forSpreadsheet(ssId).onEdit().create();
  SpreadsheetApp.getActive().toast('Awning Ruby generator trigger installed!', 'Setup Complete', 3);
}

/**
 * Backward compatibility wrappers for old function names
 */
function handleEditLeanToRuby_(e) {
  return handleEditAwningRuby_(e);
}

function installTriggerLeanToRuby_() {
  return installTriggerAwningRuby_();
}

function r_getRubyExactTemplate() { 
  return r_getLeanToRubyTemplate_(); 
}
/**
 * Copy Ruby code for the selected row to clipboard via dialog
 * Version# [12/29-10:45AM EST] by Claude Opus 4.1
 */
function copyRubyForSelectedRow_() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== 'Leads') {
    ui.alert('Please select a row in the Leads sheet');
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  if (row === 1) {
    ui.alert('Please select a data row, not the header');
    return;
  }
  
  // Get the Ruby link from column S (19)
  const rubyCell = sheet.getRange(row, 19);
  const richText = rubyCell.getRichTextValue();
  
  if (!richText) {
    ui.alert('No Ruby file found in column S for this row.\n\nMake sure columns T-AD have awning data and column AA has "Lean-to" or "A-frame".');
    return;
  }
  
  const rubyUrl = richText.getLinkUrl();
  
  if (!rubyUrl) {
    ui.alert('No Ruby link found in column S.\n\nThe cell may have text but no hyperlink.');
    return;
  }
  
  // Extract file ID from Drive URL
  const fileIdMatch = rubyUrl.match(/[-\w]{25,}/);
  if (!fileIdMatch) {
    ui.alert('Could not extract file ID from Ruby link:\n' + rubyUrl);
    return;
  }
  
  const fileId = fileIdMatch[0];
  
  try {
    // Get the Ruby file content from Drive
    const file = DriveApp.getFileById(fileId);
    const rubyCode = file.getBlob().getDataAsString();
    
    // Get display name for the dialog title
    const displayName = sheet.getRange(row, 6).getDisplayValue() || 'Row ' + row;
    
    // Create HTML dialog with copy functionality
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 15px;
              margin: 0;
            }
            h3 {
              margin-top: 0;
              color: #1a73e8;
            }
            .info {
              background: #e8f0fe;
              padding: 10px;
              border-radius: 4px;
              margin-bottom: 15px;
              font-size: 13px;
            }
            textarea {
              width: 100%;
              height: 300px;
              font-family: 'Courier New', monospace;
              font-size: 11px;
              padding: 10px;
              border: 1px solid #ccc;
              border-radius: 4px;
              resize: vertical;
            }
            .button-row {
              margin-top: 15px;
              display: flex;
              gap: 10px;
            }
            button {
              padding: 10px 24px;
              font-size: 14px;
              border: none;
              border-radius: 4px;
              cursor: pointer;
            }
            .copy-btn {
              background: #1a73e8;
              color: white;
            }
            .copy-btn:hover {
              background: #1557b0;
            }
            .close-btn {
              background: #f1f3f4;
              color: #333;
            }
            .close-btn:hover {
              background: #e0e0e0;
            }
            .success {
              color: #1e8e3e;
              font-weight: bold;
              margin-left: 15px;
              display: none;
            }
          </style>
        </head>
        <body>
          <h3>Ruby Code: ${displayName}</h3>
          <div class="info">
            <strong>Instructions:</strong> Click "Copy to Clipboard", then paste into SketchUp's Ruby Console (Window → Ruby Console)
          </div>
          <textarea id="rubyCode" readonly>${rubyCode.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</textarea>
          <div class="button-row">
            <button class="copy-btn" onclick="copyCode()">📋 Copy to Clipboard</button>
            <button class="close-btn" onclick="google.script.host.close()">Close</button>
            <span class="success" id="successMsg">✓ Copied!</span>
          </div>
          <script>
            function copyCode() {
              const textarea = document.getElementById('rubyCode');
              textarea.select();
              textarea.setSelectionRange(0, 99999);
              
              navigator.clipboard.writeText(textarea.value).then(function() {
                document.getElementById('successMsg').style.display = 'inline';
                setTimeout(function() {
                  document.getElementById('successMsg').style.display = 'none';
                }, 2000);
              }).catch(function(err) {
                // Fallback for older browsers
                document.execCommand('copy');
                document.getElementById('successMsg').style.display = 'inline';
                setTimeout(function() {
                  document.getElementById('successMsg').style.display = 'none';
                }, 2000);
              });
            }
            
            // Auto-select all text on load for easy copying
            document.getElementById('rubyCode').select();
          </script>
        </body>
      </html>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(700)
      .setHeight(500);
    
    ui.showModalDialog(htmlOutput, 'Copy Ruby Code');
    
  } catch (err) {
    ui.alert('Error loading Ruby file:\n\n' + err.message);
  }
}