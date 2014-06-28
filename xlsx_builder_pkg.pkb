CREATE OR REPLACE PACKAGE BODY "XLSX_BUILDER_PKG" 
IS

  /* Types */

  TYPE tp_XF_fmt IS RECORD
    ( numFmtId PLS_INTEGER
    , fontId PLS_INTEGER
    , fillId PLS_INTEGER
    , borderId PLS_INTEGER
    , alignment tp_alignment
    );
  TYPE tp_col_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
  TYPE tp_row_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
  TYPE tp_widths IS TABLE OF NUMBER INDEX BY PLS_INTEGER;
  
  TYPE tp_cell IS RECORD
    ( value_id NUMBER
    , style_def VARCHAR2(50)
    );

  TYPE tp_cells IS TABLE OF tp_cell INDEX BY PLS_INTEGER;
  TYPE tp_rows IS TABLE OF tp_cells INDEX BY PLS_INTEGER;

  TYPE tp_autofilter IS RECORD
    ( column_start PLS_INTEGER
    , column_end PLS_INTEGER
    , row_start PLS_INTEGER
    , row_end PLS_INTEGER
    );
  TYPE tp_autofilters IS TABLE OF tp_autofilter INDEX BY PLS_INTEGER;
  TYPE tp_hyperlink IS RECORD
    ( cell VARCHAR2(10)
    , url  VARCHAR2(1000)
    );
  TYPE tp_hyperlinks IS TABLE OF tp_hyperlink INDEX BY PLS_INTEGER;
  subtype tp_author IS VARCHAR2(32767 CHAR);
  type tp_authors is table of PLS_INTEGER index by tp_author;
  authors tp_authors;
  type tp_comment is record
    ( text VARCHAR2(32767 char)
    , author tp_author
    , row PLS_INTEGER
    , column PLS_INTEGER
    , width PLS_INTEGER
    , height PLS_INTEGER
    );
  type tp_comments is table of tp_comment index by PLS_INTEGER;
  type tp_mergecells is table of VARCHAR2(21) index by PLS_INTEGER;
  type tp_validation is record
    ( type VARCHAR2(10)
    , errorstyle VARCHAR2(32)
    , showinputmessage boolean
    , prompt VARCHAR2(32767 char)
    , title VARCHAR2(32767 char)
    , error_title VARCHAR2(32767 char)
    , error_txt VARCHAR2(32767 char)
    , showerrormessage boolean
    , formula1 VARCHAR2(32767 char)
    , formula2 VARCHAR2(32767 char)
    , allowBlank boolean
    , sqref VARCHAR2(32767 char)
    );
  type tp_validations is table of tp_validation index by PLS_INTEGER;
  type tp_sheet is record
    ( rows tp_rows
    , widths tp_widths
    , name VARCHAR2(100)
    , freeze_rows PLS_INTEGER
    , freeze_cols PLS_INTEGER
    , autofilters tp_autofilters
    , hyperlinks tp_hyperlinks
    , col_fmts tp_col_fmts
    , row_fmts tp_row_fmts
    , comments tp_comments
    , mergecells tp_mergecells
    , validations tp_validations
    );
  type tp_sheets is table of tp_sheet index by PLS_INTEGER;
  type tp_numFmt is record
    ( numFmtId PLS_INTEGER
    , formatCode VARCHAR2(100)
    );
  type tp_numFmts is table of tp_numFmt index by PLS_INTEGER;
  type tp_fill is record
    ( patternType VARCHAR2(30)
    , fgRGB VARCHAR2(8)
    );
  type tp_fills is table of tp_fill index by PLS_INTEGER;
  type tp_cellXfs is table of tp_xf_fmt index by PLS_INTEGER;
  type tp_font is record
    ( name VARCHAR2(100)
    , family PLS_INTEGER
    , fontsize NUMBER
    , theme PLS_INTEGER
    , RGB VARCHAR2(8)
    , underline boolean
    , italic boolean
    , bold boolean
    );
  type tp_fonts is table of tp_font index by PLS_INTEGER;
  type tp_border is record
    ( top VARCHAR2(17)
    , bottom VARCHAR2(17)
    , left VARCHAR2(17)
    , right VARCHAR2(17)
    );
  type tp_borders is table of tp_border index by PLS_INTEGER;
  type tp_numFmtIndexes is table of PLS_INTEGER index by PLS_INTEGER;
  type tp_strings is table of PLS_INTEGER index by VARCHAR2(32767 char);
  type tp_str_ind is table of VARCHAR2(32767 char) index by PLS_INTEGER;
  type tp_defined_name is record
    ( name VARCHAR2(32767 char)
    , ref VARCHAR2(32767 char)
    , sheet PLS_INTEGER
    );
  type tp_defined_names is table of tp_defined_name index by PLS_INTEGER;
  type tp_book is record
    ( sheets tp_sheets
    , strings tp_strings
    , str_ind tp_str_ind
    , str_cnt PLS_INTEGER := 0
    , fonts tp_fonts
    , fills tp_fills
    , borders tp_borders
    , numFmts tp_numFmts
    , cellXfs tp_cellXfs
    , numFmtIndexes tp_numFmtIndexes
    , defined_names tp_defined_names
    );

  /* Constants */

  c_local_file_header        CONSTANT RAW(4) := hextoraw( '504B0304' ); -- Local file header signature
  c_end_of_central_directory CONSTANT RAW(4) := hextoraw( '504B0506' ); -- End of central directory signature  

  /* Globals */
  workbook tp_book;
--
  FUNCTION get_workbook
    RETURN tp_book
  AS
  BEGIN
    RETURN workbook;
  END get_workbook;
  
  /* Private API */

  FUNCTION alfan_col( p_col PLS_INTEGER )
    RETURN VARCHAR2
  AS
  BEGIN
    RETURN CASE
             WHEN p_col > 702 THEN
               chr( 64 + TRUNC(( p_col - 27 ) / 676 ) ) ||
               chr( 65 + mod( TRUNC(( p_col - 1 ) / 26 ) - 1, 26 ) ) || 
               chr( 65 + mod( p_col - 1, 26 ) )
             WHEN p_col > 26 THEN
               chr( 64 + TRUNC(( p_col - 1 ) / 26 ) ) ||
               chr( 65 + mod( p_col - 1, 26 ) )
             ELSE
               chr( 64 + p_col )
           END;
  END alfan_col;

  FUNCTION col_alfan( p_col VARCHAR2 )
    RETURN PLS_INTEGER
  AS
  BEGIN
    RETURN ascii( SUBSTR( p_col, - 1 ) ) - 64 + 
                  NVL(( ascii( SUBSTR( p_col, - 2, 1 ) ) - 64 ) * 26, 0 ) +
                  NVL(( ascii( SUBSTR( p_col, - 3, 1 ) ) - 64 ) * 676,0 );
  END col_alfan;

  -- EMORKLE (2014/02/24): Moved to top, allowing usage in new_sheet
  FUNCTION add_string( p_string VARCHAR2 )
    RETURN PLS_INTEGER
  AS
    t_cnt PLS_INTEGER;
  BEGIN
    -- MKLEIN (2014/02/24): Fix to handle NULL values
    IF p_string IS NULL AND workbook.strings.count > 0 THEN
      RETURN 0;
    END IF;
    -- END Fix
    IF workbook.strings.EXISTS( p_string ) THEN
      t_cnt := workbook.strings( p_string );
    ELSE
      t_cnt := workbook.strings.count();
      workbook.str_ind( t_cnt ) := p_string;
      workbook.strings( nvl( p_string, '' ) ) := t_cnt;
    END IF;
    workbook.str_cnt := workbook.str_cnt + 1;
    RETURN t_cnt;
  END add_string;

  procedure clear_workbook
  is
    t_row_ind PLS_INTEGER;
  begin
    for s in 1 .. workbook.sheets.count()
    loop
      t_row_ind := workbook.sheets( s ).rows.first();
      while t_row_ind is not NULL
      loop
        workbook.sheets( s ).rows( t_row_ind ).delete();
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      END LOOP;
      workbook.sheets( s ).rows.delete();
      workbook.sheets( s ).widths.delete();
      workbook.sheets( s ).autofilters.delete();
      workbook.sheets( s ).hyperlinks.delete();
      workbook.sheets( s ).col_fmts.delete();
      workbook.sheets( s ).row_fmts.delete();
      workbook.sheets( s ).comments.delete();
      workbook.sheets( s ).mergecells.delete();
      workbook.sheets( s ).validations.delete();
    END LOOP;
    workbook.strings.delete();
    workbook.str_ind.delete();
    workbook.fonts.delete();
    workbook.fills.delete();
    workbook.borders.delete();
    workbook.numFmts.delete();
    workbook.cellXfs.delete();
    workbook.defined_names.delete();
    workbook := NULL;
  end;
--
  FUNCTION new_sheet( p_sheetname VARCHAR2 := NULL )
    RETURN PLS_INTEGER
  AS
    t_nr PLS_INTEGER := workbook.sheets.count() + 1;
    t_ind PLS_INTEGER;
  BEGIN
    workbook.sheets( t_nr ).NAME := nvl( dbms_xmlgen.CONVERT( translate( p_sheetname, 'a/\[]*:?', 'a' ) ), 'Sheet' || t_nr );
    IF workbook.strings.count() = 0 THEN
      workbook.str_cnt := 0;
      -- MKLEIN (2014/02/24): Insert NULL into strings on known position
      t_ind := add_string(NULL);
    END IF;
    IF workbook.fonts.count() = 0 THEN
      t_ind := get_font( 'Arial' );
    END IF;
    IF workbook.fills.count() = 0 THEN
      t_ind := get_fill( 'none' );
      t_ind := get_fill( 'gray125' );
    END IF;
    IF workbook.borders.count() = 0 THEN
      t_ind := get_border( '', '', '', '' );
    END IF;
    RETURN t_nr;
  END new_sheet;

  PROCEDURE set_col_width( p_sheet PLS_INTEGER
                         , p_col PLS_INTEGER
                         , p_format VARCHAR2
                         )
  AS
    t_width NUMBER;
    t_nr_chr PLS_INTEGER;
  BEGIN
    IF p_format IS NULL THEN
      RETURN;
    END IF;
    IF instr( p_format, ';' ) > 0 THEN
      t_nr_chr := LENGTH( TRANSLATE( SUBSTR( p_format, 1, instr( p_format, ';' ) - 1 ), 'a\"', 'a' ) );
    ELSE
      t_nr_chr := LENGTH( TRANSLATE( p_format, 'a\"', 'a' ) );
    END IF;
    t_width := TRUNC(( t_nr_chr * 7 + 5 ) / 7 * 256 ) / 256; -- assume default 11 point Calibri
    IF workbook.sheets(p_sheet).widths.EXISTS(p_col) THEN
      workbook.sheets(p_sheet).widths(p_col) := greatest(workbook.sheets(p_sheet).widths(p_col), t_width);
    ELSE
      workbook.sheets(p_sheet).widths(p_col) := greatest(t_width, 8.43);
    END IF;
  END set_col_width;


  FUNCTION OraFmt2Excel( p_format VARCHAR2 := NULL )
    RETURN VARCHAR2
  AS
    t_format VARCHAR2( 1000 ) := lower( SUBSTR( p_format, 1, 1000 ) );
  BEGIN
    t_format := REPLACE( REPLACE( REPLACE( t_format, 'HH', 'hh' ), 'hh24', 'hh' ), 'hh12', 'hh' );
    t_format := REPLACE( REPLACE( t_format, 'MI', 'mi' ), 'mi', 'mm' );
    t_format := REPLACE( REPLACE( REPLACE( t_format, 'AM', '~~' ), 'PM', '~~' ), '~~', 'AM/PM' );
    t_format := REPLACE( REPLACE( REPLACE( t_format, 'am', '~~' ), 'pm', '~~' ), '~~', 'AM/PM' );
    t_format := REPLACE( REPLACE( t_format, 'day', 'DAY' ), 'DAY', 'dddd' );
    t_format := REPLACE( REPLACE( t_format, 'dy', 'DY' ), 'DAY', 'ddd' );
    t_format := REPLACE( REPLACE( t_format, 'rr', 'RR' ), 'RR', 'YY' );
    t_format := REPLACE( REPLACE( t_format, 'month', 'MONTH' ), 'MONTH', 'mmmm' );
    t_format := REPLACE( REPLACE( t_format, 'mon', 'MON' ), 'MON', 'mmm' );
    RETURN t_format;
  END OraFmt2Excel;

  FUNCTION OraDateToExcel ( p_value IN DATE )
    RETURN NUMBER
  AS
    l_date_diff NUMBER := 0;
  BEGIN
    IF TRUNC(p_value) >= to_date('01-01-1900', 'MM-DD-YYYY') THEN
      l_date_diff := 1;
    END IF;
    RETURN ( TRUNC( p_value ) + l_date_diff ) - ( to_date('01-01-1900', 'MM-DD-YYYY') - 1 );
  END OraDateToExcel;

  FUNCTION OraNumFmt2Excel ( p_format VARCHAR2 )
    RETURN VARCHAR2
  AS
    l_mso_fmt VARCHAR2(255);
  BEGIN
    IF INSTR(p_format, 'D') > 0 THEN
      l_mso_fmt := '.' || REPLACE( substr( p_format, instr( p_format, 'D' ) + 1 ), '9', '0' );
    END IF;
    IF instr(p_format,'G') > 0 THEN
      l_mso_fmt := '#,##0' || l_mso_fmt;
    ELSE
      l_mso_fmt := '0' || l_mso_fmt;
    END IF;
    RETURN l_mso_fmt;
  END OraNumFmt2Excel;

  FUNCTION get_numFmt( p_format VARCHAR2 := NULL )
    RETURN PLS_INTEGER
  AS
    t_cnt PLS_INTEGER;
    t_numFmtId PLS_INTEGER;
  BEGIN
    IF p_format IS NULL THEN
      RETURN 0;
    END IF;
    t_cnt := workbook.numFmts.count( );
    FOR i IN 1 .. t_cnt
    LOOP
      IF workbook.numFmts( i ).formatCode = p_format THEN
        t_numFmtId := workbook.numFmts( i ).numFmtId;
        EXIT;
      END IF;
    END LOOP;
    IF t_numFmtId IS NULL THEN
      t_numFmtId := CASE
                      WHEN t_cnt = 0 THEN 164
                      ELSE workbook.numFmts( t_cnt ).numFmtId + 1
                    END;
      t_cnt := t_cnt + 1;
      workbook.numFmts( t_cnt ).numFmtId := t_numFmtId;
      workbook.numFmts( t_cnt ).formatCode := p_format;
      workbook.numFmtIndexes( t_numFmtId ) := t_cnt;
    END IF;
    RETURN t_numFmtId;
  END get_numFmt;


  FUNCTION get_font( p_name VARCHAR2
                   , p_family PLS_INTEGER := 2
                   , p_fontsize NUMBER := 8
                   , p_theme PLS_INTEGER := 1
                   , p_underline boolean := FALSE
                   , p_italic boolean := FALSE
                   , p_bold boolean := FALSE
                   , p_rgb VARCHAR2 := NULL -- this is a hex ALPHA Red Green Blue value
                   )
  RETURN PLS_INTEGER
  AS
    t_ind PLS_INTEGER;
  BEGIN
    IF workbook.fonts.count() > 0 THEN
      FOR f IN 0 .. workbook.fonts.count() - 1
      LOOP
        IF ( workbook.fonts( f ).NAME = p_name
         AND workbook.fonts( f ).family = p_family
         AND workbook.fonts( f ).fontsize = p_fontsize
         AND workbook.fonts( f ).theme = p_theme
         AND workbook.fonts( f ).underline = p_underline
         AND workbook.fonts( f ).italic = p_italic
         AND workbook.fonts( f ).bold = p_bold
         AND ( workbook.fonts( f ).rgb = p_rgb
            OR ( workbook.fonts( f ).rgb IS NULL
             AND p_rgb IS NULL 
               )
             )
           )
        THEN
          RETURN f;
        END IF;
      END LOOP;
    END IF;
    t_ind := workbook.fonts.count();
    workbook.fonts( t_ind ).name := p_name;
    workbook.fonts( t_ind ).family := p_family;
    workbook.fonts( t_ind ).fontsize := p_fontsize;
    workbook.fonts( t_ind ).theme := p_theme;
    workbook.fonts( t_ind ).underline := p_underline;
    workbook.fonts( t_ind ).italic := p_italic;
    workbook.fonts( t_ind ).bold := p_bold;
    workbook.fonts( t_ind ).rgb := p_rgb;
    RETURN t_ind;
  END get_font;

  FUNCTION get_fill( p_patternType VARCHAR2
                   , p_fgRGB       VARCHAR2 := NULL
                   )
    RETURN PLS_INTEGER
  AS
    t_ind PLS_INTEGER;
  BEGIN
    IF workbook.fills.count( ) > 0 THEN
      FOR f IN 0 .. workbook.fills.count( ) - 1
      LOOP
        IF ( workbook.fills( f ).patternType = p_patternType
         AND NVL( workbook.fills( f ).fgRGB, 'x' ) = NVL( upper( p_fgRGB ), 'x' ) )
        THEN
          RETURN f;
        END IF;
      END LOOP;
    END IF;
    t_ind := workbook.fills.count( );
    workbook.fills( t_ind ).patternType := p_patternType;
    workbook.fills( t_ind ).fgRGB := upper( p_fgRGB );
    RETURN t_ind;
  END get_fill;

  FUNCTION get_border(
      p_top    VARCHAR2 := 'thin',
      p_bottom VARCHAR2 := 'thin',
      p_left   VARCHAR2 := 'thin',
      p_right  VARCHAR2 := 'thin' )
    RETURN PLS_INTEGER
  AS
    t_ind PLS_INTEGER;
  BEGIN
    IF workbook.borders.count( ) > 0 THEN
      FOR b IN 0 .. workbook.borders.count( ) - 1
      LOOP
        IF ( NVL( workbook.borders( b ).top, 'x' ) = NVL( p_top, 'x' )
         AND NVL( workbook.borders( b ).bottom, 'x' ) = NVL( p_bottom, 'x' )
         AND NVL( workbook.borders( b ).LEFT, 'x' ) = NVL( p_left, 'x' )
         AND NVL( workbook.borders( b ).RIGHT, 'x' ) = NVL( p_right, 'x' ) )
        THEN
          RETURN b;
        END IF;
      END LOOP;
    END IF;
    t_ind := workbook.borders.count( );
    workbook.borders( t_ind ).top := p_top;
    workbook.borders( t_ind ).bottom := p_bottom;
    workbook.borders( t_ind ).left := p_left;
    workbook.borders( t_ind ).right := p_right;
    RETURN t_ind;
  END get_border;

  FUNCTION get_alignment( p_vertical   VARCHAR2 := NULL
                        , p_horizontal VARCHAR2 := NULL
                        , p_wrapText   BOOLEAN := NULL
                        )
    RETURN tp_alignment
  AS
    t_rv tp_alignment;
  BEGIN
    t_rv.vertical := p_vertical;
    t_rv.horizontal := p_horizontal;
    t_rv.wrapText := p_wrapText;
    RETURN t_rv;
  END get_alignment;

  FUNCTION get_XfId( p_sheet PLS_INTEGER
                   , p_col PLS_INTEGER
                   , p_row PLS_INTEGER
                   , p_numFmtId PLS_INTEGER := NULL
                   , p_fontId PLS_INTEGER := NULL
                   , p_fillId PLS_INTEGER := NULL
                   , p_borderId PLS_INTEGER := NULL
                   , p_alignment tp_alignment := NULL
                   )
    RETURN VARCHAR2
  AS
    t_cnt PLS_INTEGER;
    t_XfId PLS_INTEGER;
    t_XF tp_XF_fmt;
    t_col_XF tp_XF_fmt;
    t_row_XF tp_XF_fmt;
  BEGIN
    IF workbook.sheets( p_sheet ).col_fmts.exists( p_col ) THEN
      t_col_XF := workbook.sheets( p_sheet ).col_fmts( p_col );
    END IF;
    IF workbook.sheets( p_sheet ).row_fmts.exists( p_row ) THEN
      t_row_XF := workbook.sheets( p_sheet ).row_fmts( p_row );
    END IF;
    
    t_XF.numFmtId := COALESCE( p_numFmtId, t_col_XF.numFmtId, t_row_XF.numFmtId, 0 );
    t_XF.fontId   := COALESCE( p_fontId, t_col_XF.fontId, t_row_XF.fontId, 0 );
    t_XF.fillId   := COALESCE( p_fillId, t_col_XF.fillId, t_row_XF.fillId, 0 );
    t_XF.borderId := COALESCE( p_borderId, t_col_XF.borderId, t_row_XF.borderId, 0 );
    t_XF.alignment := COALESCE( p_alignment, t_col_XF.alignment, t_row_XF.alignment );
    
    IF ( t_XF.numFmtId + t_XF.fontId + t_XF.fillId + t_XF.borderId = 0
     AND t_XF.alignment.vertical IS NULL
     AND t_XF.alignment.horizontal IS NULL
     AND NOT NVL( t_XF.alignment.wrapText, FALSE ) )
    THEN
      RETURN '';
    END IF;

    IF t_XF.numFmtId > 0 THEN
      set_col_width( p_sheet, p_col, workbook.numFmts( workbook.numFmtIndexes( t_XF.numFmtId ) ).formatCode );
    END IF;
    t_cnt := workbook.cellXfs.count( );
    FOR i IN 1 .. t_cnt
    LOOP
      IF ( workbook.cellXfs( i ).numFmtId = t_XF.numFmtId
       AND workbook.cellXfs( i ).fontId = t_XF.fontId
       AND workbook.cellXfs( i ).fillId = t_XF.fillId
       AND workbook.cellXfs( i ).borderId = t_XF.borderId
       AND NVL( workbook.cellXfs( i ).alignment.vertical, 'x' ) = NVL( t_XF.alignment.vertical, 'x' )
       AND NVL( workbook.cellXfs( i ).alignment.horizontal, 'x' ) = NVL( t_XF.alignment.horizontal, 'x' )
       AND NVL( workbook.cellXfs( i ).alignment.wrapText, FALSE ) = NVL( t_XF.alignment.wrapText, FALSE ) )
      THEN
        t_XfId := i;
        EXIT;
      END IF;
    END LOOP;
    
    IF t_XfId IS NULL THEN
      t_cnt := t_cnt + 1;
      t_XfId := t_cnt;
      workbook.cellXfs( t_cnt ) := t_XF;
    END IF;
    RETURN 's="' || t_XfId || '"';
  END get_XfId;

  PROCEDURE cell( p_col PLS_INTEGER
                , p_row PLS_INTEGER
                , p_value NUMBER
                , p_numFmtId PLS_INTEGER := NULL
                , p_fontId PLS_INTEGER := NULL
                , p_fillId PLS_INTEGER := NULL
                , p_borderId PLS_INTEGER := NULL
                , p_alignment tp_alignment := NULL
                , p_sheet PLS_INTEGER := NULL
                )
  AS
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).value_id := p_value;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style_def := NULL;
    workbook.sheets( t_sheet ).ROWS( p_row )( p_col ).style_def := get_XfId( t_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment );
  END cell;

  PROCEDURE cell( p_col PLS_INTEGER
                , p_row PLS_INTEGER
                , p_value VARCHAR2
                , p_numFmtId PLS_INTEGER := NULL
                , p_fontId PLS_INTEGER := NULL
                , p_fillId PLS_INTEGER := NULL
                , p_borderId PLS_INTEGER := NULL
                , p_alignment tp_alignment := NULL
                , p_sheet PLS_INTEGER := NULL
                )
  AS
    t_sheet PLS_INTEGER := NVL( p_sheet, workbook.sheets.count( ) );
    t_alignment tp_alignment := p_alignment;
  BEGIN
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).value_id := add_string( p_value );
    IF t_alignment.wrapText IS NULL AND instr( p_value, chr( 13 ) ) > 0
    THEN
      t_alignment.wrapText := true;
    END IF;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style_def := 't="s" ' || get_XfId( t_sheet, p_col, p_row, p_numFmtId, p_fontId, p_fillId, p_borderId, t_alignment );
  END cell;

  PROCEDURE cell( p_col PLS_INTEGER
                , p_row PLS_INTEGER
                , p_value DATE
                , p_numFmtId PLS_INTEGER := NULL
                , p_fontId PLS_INTEGER := NULL
                , p_fillId PLS_INTEGER := NULL
                , p_borderId PLS_INTEGER := NULL
                , p_alignment tp_alignment := NULL
                , p_sheet PLS_INTEGER := NULL
                )
  AS
    t_numFmtId PLS_INTEGER := p_numFmtId;
    t_sheet PLS_INTEGER := NVL( p_sheet, workbook.sheets.count( ) );
  BEGIN
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).value_id := OraDatetoExcel( p_value );
    IF t_numFmtId IS NULL
     AND NOT ( workbook.sheets( t_sheet ).col_fmts.EXISTS( p_col )
           AND workbook.sheets( t_sheet ).col_fmts( p_col ).numFmtId IS NOT NULL
             )
     AND NOT ( workbook.sheets( t_sheet ).row_fmts.EXISTS( p_row )
           AND workbook.sheets( t_sheet ).row_fmts( p_row ).numFmtId IS NOT NULL
             )
    THEN
      t_numFmtId := get_numFmt( 'dd/mm/yyyy' );
    END IF;
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style_def := get_XfId( t_sheet, p_col, p_row, t_numFmtId, p_fontId, p_fillId, p_borderId, p_alignment );
  END cell;

  PROCEDURE hyperlink( p_col PLS_INTEGER
                     , p_row PLS_INTEGER
                     , p_url VARCHAR2
                     , p_value VARCHAR2 := NULL
                     , p_sheet PLS_INTEGER := NULL
                     )
  AS
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL( p_sheet, workbook.sheets.count( ) );
  BEGIN
    workbook.sheets( t_sheet ).ROWS( p_row )( p_col ).value_id := add_string( NVL( p_value, p_url ) );
    workbook.sheets( t_sheet ).rows( p_row )( p_col ).style_def := 't="s" ' || get_XfId( t_sheet, p_col, p_row, '', get_font( 'Calibri', p_theme => 10, p_underline => true ) );
    t_ind := workbook.sheets( t_sheet ).hyperlinks.count( ) + 1;
    workbook.sheets( t_sheet ).hyperlinks( t_ind ).cell := alfan_col( p_col ) || p_row;
    workbook.sheets( t_sheet ).hyperlinks( t_ind ).url := p_url;
  END hyperlink;

  PROCEDURE comment
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_text VARCHAR2
    , p_author VARCHAR2 := NULL
    , p_width PLS_INTEGER := 150
    , p_height PLS_INTEGER := 100
    , p_sheet PLS_INTEGER := NULL
    )
  AS
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    t_ind := workbook.sheets( t_sheet ).comments.count() + 1;
    workbook.sheets( t_sheet ).comments( t_ind ).row := p_row;
    workbook.sheets( t_sheet ).comments( t_ind ).column := p_col;
    workbook.sheets( t_sheet ).comments( t_ind ).text := dbms_xmlgen.convert( p_text );
    workbook.sheets( t_sheet ).comments( t_ind ).author := dbms_xmlgen.convert( p_author );
    workbook.sheets( t_sheet ).comments( t_ind ).width := p_width;
    workbook.sheets( t_sheet ).comments( t_ind ).height := p_height;
  END comment;

  PROCEDURE mergecells( p_tl_col PLS_INTEGER -- top left
                      , p_tl_row PLS_INTEGER
                      , p_br_col PLS_INTEGER -- bottom right
                      , p_br_row PLS_INTEGER
                      , p_sheet PLS_INTEGER := NULL
                      )
  AS
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL( p_sheet, workbook.sheets.count( ) );
  BEGIN
    t_ind := workbook.sheets( t_sheet ).mergecells.count( ) + 1;
    workbook.sheets( t_sheet ).mergecells( t_ind ) := alfan_col( p_tl_col ) || p_tl_row || ':' || alfan_col( p_br_col ) || p_br_row;
  END mergecells;

  PROCEDURE add_validation( p_type  VARCHAR2
                          , p_sqref VARCHAR2
                          , p_style VARCHAR2 := 'stop' -- stop, warning, information
                          , p_formula1 VARCHAR2 := NULL
                          , p_formula2 VARCHAR2 := NULL
                          , p_title VARCHAR2 := NULL
                          , p_prompt VARCHAR2 := NULL
                          , p_show_error  BOOLEAN := FALSE
                          , p_error_title VARCHAR2 := NULL
                          , p_error_txt   VARCHAR2 := NULL
                          , p_sheet PLS_INTEGER := NULL
                          )
  AS
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := NVL( p_sheet, workbook.sheets.count( ) );
  BEGIN
    t_ind := workbook.sheets( t_sheet ).validations.count( ) + 1;
    workbook.sheets( t_sheet ).validations( t_ind ).type := p_type;
    workbook.sheets( t_sheet ).validations( t_ind ).errorstyle := p_style;
    workbook.sheets( t_sheet ).validations( t_ind ).sqref := p_sqref;
    workbook.sheets( t_sheet ).validations( t_ind ).formula1 := p_formula1;
    workbook.sheets( t_sheet ).validations( t_ind ).error_title := p_error_title;
    workbook.sheets( t_sheet ).validations( t_ind ).error_txt := p_error_txt;
    workbook.sheets( t_sheet ).validations( t_ind ).title := p_title;
    workbook.sheets( t_sheet ).validations( t_ind ).prompt := p_prompt;
    workbook.sheets( t_sheet ).validations( t_ind ).showerrormessage := p_show_error;
  END add_validation;

  PROCEDURE list_validation( p_sqref_col PLS_INTEGER
                           , p_sqref_row PLS_INTEGER
                           , p_tl_col PLS_INTEGER -- top left
                           , p_tl_row PLS_INTEGER
                           , p_br_col PLS_INTEGER -- bottom right
                           , p_br_row PLS_INTEGER
                           , p_style VARCHAR2 := 'stop' -- stop, warning, information
                           , p_title VARCHAR2 := NULL
                           , p_prompt VARCHAR2 := NULL
                           , p_show_error BOOLEAN := FALSE
                           , p_error_title VARCHAR2 := NULL
                           , p_error_txt VARCHAR2 := NULL
                           , p_sheet PLS_INTEGER := NULL
                           )
  AS
  BEGIN
    add_validation( 'list'
                  , alfan_col( p_sqref_col ) || p_sqref_row
                  , p_style => lower( p_style )
                  , p_formula1 => '$' || alfan_col( p_tl_col ) || '$' || p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row
                  , p_title => p_title
                  , p_prompt => p_prompt
                  , p_show_error => p_show_error
                  , p_error_title => p_error_title
                  , p_error_txt => p_error_txt
                  , p_sheet => p_sheet
                  );
  END list_validation;

  PROCEDURE list_validation( p_sqref_col PLS_INTEGER
                           , p_sqref_row PLS_INTEGER
                           , p_defined_name VARCHAR2
                           , p_style VARCHAR2 := 'stop' -- stop, warning, information
                           , p_title VARCHAR2 := NULL
                           , p_prompt VARCHAR2 := NULL
                           , p_show_error boolean := FALSE
                           , p_error_title VARCHAR2 := NULL
                           , p_error_txt VARCHAR2 := NULL
                           , p_sheet PLS_INTEGER := NULL
                           )
  AS
  BEGIN
    add_validation( 'list'
                  , alfan_col( p_sqref_col ) || p_sqref_row
                  , p_style => lower( p_style )
                  , p_formula1 => p_defined_name 
                  , p_title => p_title
                  , p_prompt => p_prompt
                  , p_show_error => p_show_error
                  , p_error_title => p_error_title
                  , p_error_txt => p_error_txt
                  , p_sheet => p_sheet
                  ); 
  END list_validation;

  PROCEDURE defined_name( p_tl_col PLS_INTEGER -- top left
                        , p_tl_row PLS_INTEGER
                        , p_br_col PLS_INTEGER -- bottom right
                        , p_br_row PLS_INTEGER
                        , p_name VARCHAR2
                        , p_sheet PLS_INTEGER := NULL
                        , p_localsheet PLS_INTEGER := NULL
                        )
  AS
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    t_ind := workbook.defined_names.count() + 1;
    workbook.defined_names( t_ind ).name := p_name;
    workbook.defined_names( t_ind ).ref := 'Sheet' || t_sheet || '!$' || alfan_col( p_tl_col ) || '$' ||  p_tl_row || ':$' || alfan_col( p_br_col ) || '$' || p_br_row;
    workbook.defined_names( t_ind ).sheet := p_localsheet;
  END defined_name;

  PROCEDURE set_column_width( p_col PLS_INTEGER
                            , p_width NUMBER
                            , p_sheet PLS_INTEGER := NULL
                            )
  AS
  BEGIN
    workbook.sheets( nvl( p_sheet, workbook.sheets.count() ) ).widths( p_col ) := p_width;
  END set_column_width;

  PROCEDURE set_column( p_col PLS_INTEGER
                      , p_numFmtId PLS_INTEGER := NULL
                      , p_fontId PLS_INTEGER := NULL
                      , p_fillId PLS_INTEGER := NULL
                      , p_borderId PLS_INTEGER := NULL
                      , p_alignment tp_alignment := NULL
                      , p_sheet PLS_INTEGER := NULL
                      )
  AS
    t_sheet PLS_INTEGER := NVL( p_sheet, workbook.sheets.count( ) );
  BEGIN
    workbook.sheets( t_sheet ).col_fmts( p_col ).numFmtId := p_numFmtId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).fontId := p_fontId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).fillId := p_fillId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).borderId := p_borderId;
    workbook.sheets( t_sheet ).col_fmts( p_col ).alignment := p_alignment;
  END set_column;

  PROCEDURE set_row( p_row PLS_INTEGER
                   , p_numFmtId PLS_INTEGER := NULL
                   , p_fontId PLS_INTEGER := NULL
                   , p_fillId PLS_INTEGER := NULL
                   , p_borderId PLS_INTEGER := NULL
                   , p_alignment tp_alignment := NULL
                   , p_sheet PLS_INTEGER := NULL
                   )
  AS
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    workbook.sheets( t_sheet ).row_fmts( p_row ).numFmtId := p_numFmtId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).fontId := p_fontId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).fillId := p_fillId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).borderId := p_borderId;
    workbook.sheets( t_sheet ).row_fmts( p_row ).alignment := p_alignment;
  END set_row;

  PROCEDURE freeze_rows( p_nr_rows PLS_INTEGER := 1
                       , p_sheet PLS_INTEGER := NULL
                       )
  AS
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    workbook.sheets( t_sheet ).freeze_cols := NULL;
    workbook.sheets( t_sheet ).freeze_rows := p_nr_rows;
  END freeze_rows;

  PROCEDURE freeze_cols( p_nr_cols PLS_INTEGER := 1
                       , p_sheet PLS_INTEGER := NULL
                       )
  AS
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    workbook.sheets( t_sheet ).freeze_rows := NULL;
    workbook.sheets( t_sheet ).freeze_cols := p_nr_cols;
  END freeze_cols;

  PROCEDURE freeze_pane( p_col PLS_INTEGER
                       , p_row PLS_INTEGER
                       , p_sheet PLS_INTEGER := NULL
                       )
  AS
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    workbook.sheets( t_sheet ).freeze_rows := p_row;
    workbook.sheets( t_sheet ).freeze_cols := p_col;
  END freeze_pane;

  PROCEDURE set_autofilter( p_column_start PLS_INTEGER := NULL
                          , p_column_end PLS_INTEGER := NULL
                          , p_row_start PLS_INTEGER := NULL
                          , p_row_end PLS_INTEGER := NULL
                          , p_sheet PLS_INTEGER := NULL
                          )
  AS
    t_ind PLS_INTEGER;
    t_sheet PLS_INTEGER := nvl( p_sheet, workbook.sheets.count() );
  BEGIN
    t_ind := 1;
    workbook.sheets( t_sheet ).autofilters( t_ind ).column_start := p_column_start;
    workbook.sheets( t_sheet ).autofilters( t_ind ).column_end := p_column_end;
    workbook.sheets( t_sheet ).autofilters( t_ind ).row_start := p_row_start;
    workbook.sheets( t_sheet ).autofilters( t_ind ).row_end := p_row_end;
    defined_name( p_column_start
                , p_row_start
                , p_column_end
                , p_row_end
                , '_xlnm._FilterDatabase'
                , t_sheet
                , t_sheet - 1
                );
  END set_autofilter;

  FUNCTION finish
    RETURN BLOB
  AS
    t_excel BLOB;
    t_xxx CLOB;
    t_tmp VARCHAR2(32767 CHAR);
    t_str VARCHAR2(32767 CHAR);
    t_c NUMBER;
    t_h NUMBER;
    t_w NUMBER;
    t_cw NUMBER;
    t_cell VARCHAR2(1000 CHAR);
    t_row_ind PLS_INTEGER;
    t_col_min PLS_INTEGER;
    t_col_max PLS_INTEGER;
    t_col_ind PLS_INTEGER;
    t_len PLS_INTEGER;
  BEGIN
    dbms_lob.createtemporary( t_excel, true );
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';

    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '<Override PartName="/xl/worksheets/sheet' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    END LOOP;

    t_xxx := t_xxx || '
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';

    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      IF workbook.sheets( s ).comments.count() > 0
      THEN
        t_xxx := t_xxx || '<Override PartName="/xl/comments' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>';
      END IF;
    END LOOP;

    t_xxx := t_xxx || '</Types>';
    zip_util_pkg.add_file( t_excel, '[Content_Types].xml', t_xxx );

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>' || sys_context( 'userenv', 'os_user' ) || '</dc:creator>
<cp:lastModifiedBy>' || sys_context( 'userenv', 'os_user' ) || '</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">' || to_char( current_timestamp, 'yyyy-mm-dd"T"hh24:mi:ssTZH:TZM' ) || '</dcterms:modified>
</cp:coreProperties>';
    zip_util_pkg.add_file( t_excel, 'docProps/core.xml', t_xxx );
    
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>FALSE</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
<vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>' || workbook.sheets.count() || '</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="' || workbook.sheets.count() || '" baseType="lpstr">';

    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '<vt:lpstr>' || workbook.sheets( s ).NAME || '</vt:lpstr>';
    END LOOP;
    
    t_xxx := t_xxx || '</vt:vector>
</TitlesOfParts>
<LinksUpToDate>FALSE</LinksUpToDate>
<SharedDoc>FALSE</SharedDoc>
<HyperlinksChanged>FALSE</HyperlinksChanged>
<AppVersion>14.0300</AppVersion>
</Properties>';
    zip_util_pkg.add_file( t_excel, 'docProps/app.xml', t_xxx );

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
    zip_util_pkg.add_file( t_excel, '_rels/.rels', t_xxx );

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">';
    IF workbook.numFmts.count() > 0
    THEN
      t_xxx := t_xxx || '<numFmts count="' || workbook.numFmts.count() || '">';
      FOR n IN 1 .. workbook.numFmts.count()
      LOOP
        t_xxx := t_xxx || '<numFmt numFmtId="' || workbook.numFmts( n ).numFmtId || '" formatCode="' || workbook.numFmts( n ).formatCode || '"/>';
      END LOOP;
      t_xxx := t_xxx || '</numFmts>';
    END IF;
    t_xxx := t_xxx || '<fonts count="' || workbook.fonts.count() || '" x14ac:knownFonts="1">';
    FOR f IN 0 .. workbook.fonts.count() - 1
    LOOP
      t_xxx := t_xxx || '<font>' || 
               CASE WHEN workbook.fonts( f ).bold THEN '<b/>' END ||
               CASE WHEN workbook.fonts( f ).italic THEN '<i/>' END ||
               CASE WHEN workbook.fonts( f ).underline THEN '<u/>' END ||
               '<sz val="' || to_char( workbook.fonts( f ).fontsize, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' )  || '"/>
               <color ' || CASE WHEN workbook.fonts( f ).rgb IS NOT NULL THEN 'rgb="' || workbook.fonts( f ).rgb ELSE 'theme="' || workbook.fonts( f ).theme END || '"/>
               <name val="' || workbook.fonts( f ).NAME || '"/>
               <family val="' || workbook.fonts( f ).family || '"/>
               <scheme val="none"/>
               </font>';
    END LOOP;
    t_xxx := t_xxx || '</fonts>
<fills count="' || workbook.fills.count() || '">';
    for f in 0 .. workbook.fills.count() - 1
    loop
      t_xxx := t_xxx || '<fill><patternFill patternType="' || workbook.fills( f ).patternType || '">' ||
         case when workbook.fills( f ).fgRGB is not NULL then '<fgColor rgb="' || workbook.fills( f ).fgRGB || '"/>' end ||
         '</patternFill></fill>';
    END LOOP;
    t_xxx := t_xxx || '</fills>
<borders count="' || workbook.borders.count() || '">';
    FOR b IN 0 .. workbook.borders.count() - 1
    LOOP
      t_xxx := t_xxx || '<border>' ||
               CASE WHEN workbook.borders( b ).LEFT   IS NULL THEN '<left/>'   ELSE '<left style="'   || workbook.borders( b ).LEFT   || '"/>' END ||
               CASE WHEN workbook.borders( b ).RIGHT  IS NULL THEN '<right/>'  ELSE '<right style="'  || workbook.borders( b ).RIGHT  || '"/>' END ||
               CASE WHEN workbook.borders( b ).top    IS NULL THEN '<top/>'    ELSE '<top style="'    || workbook.borders( b ).top    || '"/>' END ||
               CASE WHEN workbook.borders( b ).bottom IS NULL THEN '<bottom/>' ELSE '<bottom style="' || workbook.borders( b ).bottom || '"/>' END ||
               '</border>';
    END LOOP;
    t_xxx := t_xxx || '</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="' || ( workbook.cellXfs.count() + 1 ) || '">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    FOR x IN 1 .. workbook.cellXfs.count()
    LOOP
      t_xxx := t_xxx || '<xf numFmtId="' || workbook.cellXfs( x ).numFmtId ||
               '" fontId="' || workbook.cellXfs( x ).fontId || 
               '" fillId="' || workbook.cellXfs( x ).fillId || 
               '" borderId="' || workbook.cellXfs( x ).borderId || '">';
      IF (  workbook.cellXfs( x ).alignment.horizontal IS NOT NULL
         OR workbook.cellXfs( x ).alignment.vertical IS NOT NULL
         OR workbook.cellXfs( x ).alignment.wrapText
         )
      THEN
        t_xxx := t_xxx || '<alignment' ||
                 CASE WHEN workbook.cellXfs( x ).alignment.horizontal IS NOT NULL THEN ' horizontal="' || workbook.cellXfs( x ).alignment.horizontal || '"' END ||
                 CASE WHEN workbook.cellXfs( x ).alignment.vertical IS NOT NULL THEN ' vertical="' || workbook.cellXfs( x ).alignment.vertical || '"' END ||
                 CASE WHEN workbook.cellXfs( x ).alignment.wrapText THEN ' wrapText="true"' END || '/>';
      END IF;
      t_xxx := t_xxx || '</xf>';
    END LOOP;
    t_xxx := t_xxx || '</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
<extLst>
<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
<x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
</ext>
</extLst>
</styleSheet>';
    zip_util_pkg.add_file( t_excel, 'xl/styles.xml', t_xxx );

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/>
<workbookPr date1904="FALSE" defaultThemeVersion="124226"/>
<bookViews>
<workbookView xWindow="120" yWindow="45" windowWidth="19155" windowHeight="4935"/>
</bookViews>
<sheets>';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '<sheet name="' || workbook.sheets( s ).name || '" sheetId="' || s || '" r:id="rId' || ( 9 + s ) || '"/>';
    END LOOP;
    t_xxx := t_xxx || '</sheets>';
    IF workbook.defined_names.count() > 0
    THEN
      t_xxx := t_xxx || '<definedNames>';
      FOR s IN 1 .. workbook.defined_names.count()
      LOOP
        t_xxx := t_xxx || '<definedName name="' || workbook.defined_names( s ).NAME || '"' ||
                 CASE WHEN workbook.defined_names( s ).sheet IS NOT NULL THEN ' localSheetId="' || to_char( workbook.defined_names( s ).sheet ) || '"' END ||
                 '>' || workbook.defined_names( s ).ref || '</definedName>';
      END LOOP;
      t_xxx := t_xxx || '</definedNames>';
    END IF;
    t_xxx := t_xxx || '<calcPr calcId="144525"/></workbook>';
    zip_util_pkg.add_file( t_excel, 'xl/workbook.xml', t_xxx );
    
    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="MS P????"/>
<a:font script="Hang" typeface="?? ??"/>
<a:font script="Hans" typeface="??"/>
<a:font script="Hant" typeface="????"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
<a:font script="Geor" typeface="Sylfaen"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:shade val="51000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="80000">
<a:schemeClr val="phClr">
<a:shade val="93000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="94000"/>
<a:satMod val="135000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>';
    zip_util_pkg.add_file( t_excel, 'xl/theme/theme1.xml', t_xxx );

    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_col_min := 16384;
      t_col_max := 1;
      t_row_ind := workbook.sheets( s ).rows.first();
      WHILE t_row_ind IS NOT NULL
      LOOP
        t_col_min := least( t_col_min, workbook.sheets( s ).rows( t_row_ind ).first() );
        t_col_max := greatest( t_col_max, workbook.sheets( s ).rows( t_row_ind ).last() );
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      END LOOP;
      t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
<dimension ref="' || alfan_col( t_col_min ) || workbook.sheets( s ).rows.first() || ':' || alfan_col( t_col_max ) || workbook.sheets( s ).rows.last() || '"/>
<sheetViews>
<sheetView' || CASE WHEN s = 1 THEN ' tabSelected="1"' END || ' workbookViewId="0">';
      IF workbook.sheets( s ).freeze_rows > 0 AND workbook.sheets( s ).freeze_cols > 0
      THEN
        t_xxx := t_xxx || 
               ( '<pane xSplit="' || workbook.sheets( s ).freeze_cols || '" ' || 
                 'ySplit="' || workbook.sheets( s ).freeze_rows || '" ' ||
                 'topLeftCell="' || alfan_col( workbook.sheets( s ).freeze_cols + 1 ) || ( workbook.sheets( s ).freeze_rows + 1 ) || '" ' ||
                 'activePane="bottomLeft" state="frozen"/>'
               );
      ELSE
        IF workbook.sheets( s ).freeze_rows > 0
        THEN
          t_xxx := t_xxx || '<pane ySplit="' || workbook.sheets( s ).freeze_rows || '" topLeftCell="A' || ( workbook.sheets( s ).freeze_rows + 1 ) || '" activePane="bottomLeft" state="frozen"/>';
        END IF;
        IF workbook.sheets( s ).freeze_cols > 0
        THEN
          t_xxx := t_xxx || '<pane xSplit="' || workbook.sheets( s ).freeze_cols || '" topLeftCell="' || alfan_col( workbook.sheets( s ).freeze_cols + 1 ) || '1" activePane="bottomLeft" state="frozen"/>';
        END IF;
      END IF;
      t_xxx := t_xxx || '</sheetView>
</sheetViews>
<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>';
      IF workbook.sheets( s ).widths.count() > 0
      THEN
        t_xxx := t_xxx || '<cols>';
        t_col_ind := workbook.sheets( s ).widths.first();
        WHILE t_col_ind IS NOT NULL
        LOOP
          t_xxx := t_xxx ||
                   '<col min="' || t_col_ind || 
                   '" max="' || t_col_ind || 
                   '" width="' || to_char( workbook.sheets( s ).widths( t_col_ind ), 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) || 
                   '" customWidth="1"/>';
          t_col_ind := workbook.sheets( s ).widths.next( t_col_ind );
        END LOOP;
        t_xxx := t_xxx || '</cols>';
      END IF;
      t_xxx := t_xxx || '<sheetData>';
      t_row_ind := workbook.sheets( s ).rows.first();
      t_tmp := NULL;
      WHILE t_row_ind IS NOT NULL
      LOOP
        t_tmp :=  t_tmp || '<row r="' || t_row_ind || '" spans="' || t_col_min || ':' || t_col_max || '">';
        t_len := length( t_tmp );
        t_col_ind := workbook.sheets( s ).rows( t_row_ind ).first();
        WHILE t_col_ind IS NOT NULL
        LOOP
          t_cell := '<c r="' || alfan_col( t_col_ind ) || t_row_ind || '"' ||
                    ' ' || workbook.sheets( s ).ROWS( t_row_ind )( t_col_ind ).style_def ||
                    '><v>' || to_char( workbook.sheets( s ).ROWS( t_row_ind )( t_col_ind ).value_id, 'TM9', 'NLS_NUMERIC_CHARACTERS=.,' ) ||
                    '</v></c>';
          IF t_len > 32000
          THEN
            dbms_lob.writeappend( t_xxx, t_len, t_tmp );
            t_tmp := NULL;
            t_len := 0;
          END IF;
          t_tmp :=  t_tmp || t_cell;
          t_len := t_len + length( t_cell );
          t_col_ind := workbook.sheets( s ).rows( t_row_ind ).next( t_col_ind );
        END LOOP;
        t_tmp :=  t_tmp || '</row>';
        t_row_ind := workbook.sheets( s ).rows.next( t_row_ind );
      END LOOP;
      t_tmp :=  t_tmp || '</sheetData>';
      t_len := length( t_tmp );
      dbms_lob.writeappend( t_xxx, t_len, t_tmp );
      FOR A IN 1 ..  workbook.sheets( s ).autofilters.count()
      LOOP
        t_xxx := t_xxx || '<autoFilter ref="' ||
                 alfan_col( nvl( workbook.sheets( s ).autofilters( A ).column_start, t_col_min ) ) ||
                 nvl( workbook.sheets( s ).autofilters( a ).row_start, workbook.sheets( s ).rows.first() ) || ':' ||
                 alfan_col( COALESCE( workbook.sheets( s ).autofilters( A ).column_end, workbook.sheets( s ).autofilters( A ).column_start, t_col_max ) ) ||
                 nvl( workbook.sheets( s ).autofilters( a ).row_end, workbook.sheets( s ).rows.last() ) || '"/>';
      END LOOP;
      IF workbook.sheets( s ).mergecells.count() > 0
      THEN
        t_xxx := t_xxx || '<mergeCells count="' || to_char( workbook.sheets( s ).mergecells.count() ) || '">';
        FOR m IN 1 ..  workbook.sheets( s ).mergecells.count()
        LOOP
          t_xxx := t_xxx || '<mergeCell ref="' || workbook.sheets( s ).mergecells( m ) || '"/>';
        END LOOP;
        t_xxx := t_xxx || '</mergeCells>';
      END IF;

      IF workbook.sheets( s ).validations.count() > 0
      THEN
        t_xxx := t_xxx || '<dataValidations count="' || to_char( workbook.sheets( s ).validations.count() ) || '">';
        FOR m IN 1 ..  workbook.sheets( s ).validations.count()
        LOOP
          t_xxx := t_xxx || '<dataValidation' ||
                   ' type="' || workbook.sheets( s ).validations( m ).TYPE || '"' ||
                   ' errorStyle="' || workbook.sheets( s ).validations( m ).errorstyle || '"' ||
                   ' allowBlank="' || CASE WHEN nvl( workbook.sheets( s ).validations( m ).allowBlank, TRUE ) THEN '1' ELSE '0' END || '"' ||
                   ' sqref="' || workbook.sheets( s ).validations( m ).sqref || '"';
          IF workbook.sheets( s ).validations( m ).prompt IS NOT NULL
          THEN
            t_xxx := t_xxx || ' showInputMessage="1" prompt="' || workbook.sheets( s ).validations( m ).prompt || '"';
            IF workbook.sheets( s ).validations( m ).title IS NOT NULL
            THEN
              t_xxx := t_xxx || ' promptTitle="' || workbook.sheets( s ).validations( m ).title || '"';
            END IF;
          END IF;
          IF workbook.sheets( s ).validations( m ).showerrormessage
          THEN
            t_xxx := t_xxx || ' showErrorMessage="1"';
            IF workbook.sheets( s ).validations( m ).error_title IS NOT NULL
            THEN
              t_xxx := t_xxx || ' errorTitle="' || workbook.sheets( s ).validations( m ).error_title || '"';
            END IF;
            if workbook.sheets( s ).validations( m ).error_txt IS NOT NULL
            then
              t_xxx := t_xxx || ' error="' || workbook.sheets( s ).validations( m ).error_txt || '"';
            END IF;
          END IF;
          t_xxx := t_xxx || '>';
          IF workbook.sheets( s ).validations( m ).formula1 IS NOT NULL
          then
            t_xxx := t_xxx || '<formula1>' || workbook.sheets( s ).validations( m ).formula1 || '</formula1>';
          END IF;
          IF workbook.sheets( s ).validations( m ).formula2 IS NOT NULL
          THEN
            t_xxx := t_xxx || '<formula2>' || workbook.sheets( s ).validations( m ).formula2 || '</formula2>';
          END IF;
          t_xxx := t_xxx || '</dataValidation>';
        END LOOP;
        t_xxx := t_xxx || '</dataValidations>';
      END IF;

      IF workbook.sheets( s ).hyperlinks.count() > 0
      THEN
        t_xxx := t_xxx || '<hyperlinks>';
        FOR h IN 1 ..  workbook.sheets( s ).hyperlinks.count()
        LOOP
          t_xxx := t_xxx || '<hyperlink ref="' || workbook.sheets( s ).hyperlinks( h ).cell || '" r:id="rId' || h || '"/>';
        END LOOP;
        t_xxx := t_xxx || '</hyperlinks>';
      END IF;
      t_xxx := t_xxx || '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      IF workbook.sheets( s ).comments.count() > 0
      THEN
        t_xxx := t_xxx || '<legacyDrawing r:id="rId' || ( workbook.sheets( s ).hyperlinks.count() + 1 ) || '"/>';
      END IF;

      t_xxx := t_xxx || '</worksheet>';
      zip_util_pkg.add_file( t_excel, 'xl/worksheets/sheet' || s || '.xml', t_xxx );

      IF workbook.sheets( s ).hyperlinks.count() > 0 OR workbook.sheets( s ).comments.count() > 0
      THEN
        t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        IF workbook.sheets( s ).comments.count() > 0
        THEN
          t_xxx := t_xxx || '<Relationship Id="rId' || ( workbook.sheets( s ).hyperlinks.count() + 2 ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' || s || '.xml"/>';
          t_xxx := t_xxx || '<Relationship Id="rId' || ( workbook.sheets( s ).hyperlinks.count() + 1 ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || s || '.vml"/>';
        END IF;
        FOR h IN 1 ..  workbook.sheets( s ).hyperlinks.count()
        LOOP
          t_xxx := t_xxx || '<Relationship Id="rId' || h || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' || workbook.sheets( s ).hyperlinks( h ).url || '" TargetMode="External"/>';
        END LOOP;
        t_xxx := t_xxx || '</Relationships>';
        zip_util_pkg.add_file( t_excel, 'xl/worksheets/_rels/sheet' || s || '.xml.rels', t_xxx );
      END IF;

      IF workbook.sheets( s ).comments.count() > 0
      THEN
        DECLARE
          cnt PLS_INTEGER;
          author_ind tp_author;
        BEGIN
          authors.delete();
          FOR c IN 1 .. workbook.sheets( s ).comments.count()
          LOOP
            authors( workbook.sheets( s ).comments( c ).author ) := 0;
          END LOOP;
          t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors>';
          cnt := 0;
          author_ind := authors.FIRST();
          WHILE author_ind IS NOT NULL OR authors.NEXT( author_ind ) IS NOT NULL
          LOOP
            authors( author_ind ) := cnt;
            t_xxx := t_xxx || '<author>' || author_ind || '</author>';
            cnt := cnt + 1;
            author_ind := authors.next( author_ind );
          END LOOP;
        END;
        t_xxx := t_xxx || '</authors><commentList>';
        FOR c IN 1 .. workbook.sheets( s ).comments.count()
        LOOP
          t_xxx := t_xxx || '<comment ref="' || alfan_col( workbook.sheets( s ).comments( c ).COLUMN ) ||
                   to_char( workbook.sheets( s ).comments( c ).ROW || 
                   '" authorId="' || authors( workbook.sheets( s ).comments( c ).author ) ) || '">
<text>';
          IF workbook.sheets( s ).comments( c ).author IS NOT NULL
          THEN
            t_xxx := t_xxx || 
                     '<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
                     workbook.sheets( s ).comments( c ).author || ':</t></r>';
          END IF;
          t_xxx := t_xxx || 
                   '<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t xml:space="preserve">' ||
                   CASE WHEN workbook.sheets( s ).comments( c ).author IS NOT NULL THEN '
' end || workbook.sheets( s ).comments( c ).text || '</t></r></text></comment>';
        END LOOP;
        t_xxx := t_xxx || '</commentList></comments>';
        zip_util_pkg.add_file( t_excel, 'xl/comments' || s || '.xml', t_xxx );

        t_xxx := '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
        FOR c IN 1 .. workbook.sheets( s ).comments.count()
        LOOP
          t_xxx := t_xxx || '<v:shape id="_x0000_s' || to_char( c ) || 
                   '" type="#_x0000_t202" style="position:absolute;margin-left:35.25pt;margin-top:3pt;z-index:' || to_char( c ) ||
                   ';visibility:hidden;" fillcolor="#ffffe1" o:insetmode="auto">
<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>
<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>
<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>';
          t_w := workbook.sheets( s ).comments( c ).width;
          t_c := 1;
          LOOP
            IF workbook.sheets( s ).widths.EXISTS( workbook.sheets( s ).comments( c ).column + t_c )
            THEN
              t_cw := 256 * workbook.sheets( s ).widths( workbook.sheets( s ).comments( c ).column + t_c ); 
              t_cw := trunc( ( t_cw + 18 ) / 256 * 7); -- assume default 11 point Calibri
            ELSE
              t_cw := 64;
            END IF;
            EXIT WHEN t_w < t_cw;
            t_c := t_c + 1;
            t_w := t_w - t_cw;
          END LOOP;
          t_h := workbook.sheets( s ).comments( c ).height;
          t_xxx := t_xxx || to_char( '<x:Anchor>' || workbook.sheets( s ).comments( c ).column || ',15,' ||
                                     workbook.sheets( s ).comments( c ).row || ',30,' ||
                                     ( workbook.sheets( s ).comments( c ).column + t_c - 1 ) || ',' || round( t_w ) || ',' ||
                                     ( workbook.sheets( s ).comments( c ).row + 1 + trunc( t_h / 20 ) ) || ',' || mod( t_h, 20 ) || 
                                     '</x:Anchor>' );
          t_xxx := t_xxx || to_char( '<x:AutoFill>FALSE</x:AutoFill><x:Row>' ||
                                     ( workbook.sheets( s ).comments( c ).row - 1 ) ||
                                     '</x:Row><x:Column>' ||
                                     ( workbook.sheets( s ).comments( c ).column - 1 ) || 
                                     '</x:Column></x:ClientData></v:shape>' );
        END LOOP;
        t_xxx := t_xxx || '</xml>';
        zip_util_pkg.add_file( t_excel, 'xl/drawings/vmlDrawing' || s || '.vml', t_xxx );
      END IF;

    END LOOP;

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>';
    FOR s IN 1 .. workbook.sheets.count()
    LOOP
      t_xxx := t_xxx || '
<Relationship Id="rId' || ( 9 + s ) || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' || s || '.xml"/>';
    END LOOP;
    t_xxx := t_xxx || '</Relationships>';
    zip_util_pkg.add_file( t_excel, 'xl/_rels/workbook.xml.rels', t_xxx );

    t_xxx := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' || workbook.str_cnt || '" uniqueCount="' || workbook.strings.count() || '">';
    t_tmp := NULL;
    FOR i IN 0 .. workbook.str_ind.count() - 1
    LOOP
      t_str := '<si><t>' || dbms_xmlgen.CONVERT( substr( workbook.str_ind( i ), 1, 32000 ) ) || '</t></si>';
      IF length( t_tmp ) + length( t_str ) > 32000
      THEN
        t_xxx := t_xxx || t_tmp;
        t_tmp := NULL;
      END IF;
      t_tmp := t_tmp || t_str;
    END LOOP;
    t_xxx := t_xxx || t_tmp || '</sst>';
    zip_util_pkg.add_file( t_excel, 'xl/sharedStrings.xml', t_xxx );
    zip_util_pkg.finish_zip( t_excel );
    clear_workbook;
    RETURN t_excel;
  END finish;

  FUNCTION query2sheet( p_sql            VARCHAR2
                      , p_column_headers BOOLEAN := TRUE
                      , p_sheet PLS_INTEGER := NULL
                      )
    RETURN BLOB
  AS
    t_sheet PLS_INTEGER;
    t_c INTEGER;
    t_col_cnt INTEGER;
    t_desc_tab dbms_sql.desc_tab2;
    d_tab dbms_sql.date_table;
    n_tab dbms_sql.number_table;
    v_tab dbms_sql.varchar2_table;
    t_bulk_size PLS_INTEGER := 200;
    t_r INTEGER;
    t_cur_row PLS_INTEGER;
  BEGIN
    t_sheet := COALESCE( p_sheet, new_sheet );
    t_c := dbms_sql.open_cursor;
    dbms_sql.parse( t_c, p_sql, dbms_sql.native );
    dbms_sql.describe_columns2( t_c, t_col_cnt, t_desc_tab );
    
    FOR c IN 1 .. t_col_cnt LOOP
      IF p_column_headers THEN
        cell( c, 1, t_desc_tab( c ).col_name, p_sheet => t_sheet );
      END IF;
      CASE
        WHEN t_desc_tab( c ).col_type IN( 2, 100, 101 ) THEN
          dbms_sql.define_array( t_c, c, n_tab, t_bulk_size, 1 );
        WHEN t_desc_tab( c ).col_type IN( 12, 178, 179, 180, 181, 231 ) THEN
          dbms_sql.define_array( t_c, c, d_tab, t_bulk_size, 1 );
        WHEN t_desc_tab( c ).col_type IN( 1, 8, 9, 96, 112 ) THEN
          dbms_sql.define_array( t_c, c, v_tab, t_bulk_size, 1 );
        ELSE
          NULL;
      END CASE;
    END LOOP;

    t_cur_row := CASE WHEN p_column_headers THEN 2 ELSE 1 END;

    t_r := dbms_sql.execute( t_c );
    LOOP
      t_r := dbms_sql.fetch_rows( t_c );
      IF t_r > 0 THEN
        FOR c IN 1 .. t_col_cnt
        LOOP
          CASE
            WHEN t_desc_tab( c ).col_type IN( 2, 100, 101 ) THEN
              dbms_sql.column_value( t_c, c, n_tab );
              FOR i IN 0 .. t_r - 1 LOOP
                IF n_tab( i + n_tab.first( ) ) IS NOT NULL THEN
                  cell( c, t_cur_row + i, n_tab( i + n_tab.first( ) ), p_sheet => t_sheet );
                END IF;
              END LOOP;
              n_tab.delete;
            WHEN t_desc_tab( c ).col_type IN( 12, 178, 179, 180, 181, 231 ) THEN
              dbms_sql.column_value( t_c, c, d_tab );
              FOR i IN 0 .. t_r - 1 LOOP
                IF d_tab( i + d_tab.first( ) ) IS NOT NULL THEN
                  cell( c, t_cur_row + i, d_tab( i + d_tab.first( ) ), p_sheet => t_sheet );
                END IF;
              END LOOP;
              d_tab.delete;
            WHEN t_desc_tab( c ).col_type IN( 1, 8, 9, 96, 112 ) THEN
              dbms_sql.column_value( t_c, c, v_tab );
              FOR i IN 0 .. t_r - 1 LOOP
                IF v_tab( i + v_tab.first( ) ) IS NOT NULL THEN
                  cell( c, t_cur_row + i, v_tab( i + v_tab.first( ) ), p_sheet => t_sheet );
                END IF;
              END LOOP;
              v_tab.delete;
            ELSE
              NULL;
          END CASE;
        END LOOP;
      END IF;
      EXIT WHEN t_r != t_bulk_size;
      t_cur_row := t_cur_row + t_r;
    END LOOP;
    dbms_sql.close_cursor( t_c );
    RETURN finish;
  EXCEPTION
  WHEN OTHERS THEN
    IF dbms_sql.is_open( t_c ) THEN
      dbms_sql.close_cursor( t_c );
    END IF;
    RETURN NULL;
  END query2sheet;
END;
/