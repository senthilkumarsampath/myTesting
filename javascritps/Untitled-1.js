// Spaltenbreitenoptimierung
// --------------------------------------------

/*
	V 1.0 vom 17.11.2021
	Voraussetzung: Einfügemarke muss in der Tabelle blinken.
*/

// ------------------------------------------------------------------------------------
// Variablen, gerne zu bearbeiten
	var zugabe = 0.2  	// Zugabe in Millimetern
// ------------------------------------------------------------------------------------

//	Variablen, intern
	var col, c = null;
	var textwidth, cwidth = 0;
	var t = null;

// Auswahl?
	if ( app.selection.length == 0 || app.selection[0].parent.parent.constructor.name != "Table") {
		alert ( "Bitte Einfügemarke in Tabelle stellen", "Keine Tabelle" );
		exit();
	}

// Tabelle ansprechen
	var t = app.selection[0].parent.parent;

// Spalten von rechts nach links durchgehen
	for ( var j = t.columns.length - 1; j >= 0 ; j-- ) {

		col = t.columns[j];

		// durch alle Zeilen durchgehen und die Textbreite feststellen
			for ( var i = 0; i < col.cells.length; i++ ) {
				c = col.cells[i];
				textwidth = c.texts.firstItem().insertionPoints.lastItem().horizontalOffset - c.texts.firstItem().insertionPoints.firstItem().horizontalOffset;
				cwidth = Math.max(cwidth, textwidth);
			}

		// der Spalte die ermittelte Maximalbreite zuweisen und um die Zugabe erhöhen	
			col.width = cwidth + zugabe;

		// und wieder auf Null
			cwidth = 0;
	}

alert ( "Fertig!", "Fertig!" );