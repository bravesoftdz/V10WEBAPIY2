{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 21/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMENTATIONCAT ()
Mots clefs ... : TOF;PGAUGMENTATIONCAT
*****************************************************************}
Unit UTofPGAugmentationCat ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTofPGAugmentationSal,
     HSysMenu,
     UTOF ; 

Type
  TOF_PGAUGMENTATIONCAT = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure TOF_PGAUGMENTATIONCAT.OnArgument (S : String ) ;
var Grille : THGrid;
   // Emploi : String;
   Titres : HTStringList;
    HMTRAD : THSystemMenu;
begin
  Inherited ;
//        Emploi := ReadTokenPipe(S,';');
        Grille := THGrid(GeTcontrol('GLIBEMPLOI'));
        Grille.ColCount := 5;
        Titres := HTStringList.Create;
        Titres.Insert(0, 'Libellé emploi');
        Titres.Insert(1, 'Nb salariés');
        Titres.Insert(2, 'Salaire min');
        Titres.Insert(3, 'Salaire Max');
        Titres.Insert(4, 'Salaire moyen');
        Grille.Titres := Titres;
        Titres.free;
        Grille.ColFormats[0] := 'CB=PGLIBEMPLOI';
        Grille.ColFormats[1] := '# ##0';
        Grille.ColFormats[2] := '# ##0.00';
        Grille.ColFormats[3] := '# ##0.00';
        Grille.ColFormats[4] := '# ##0.00';
        Grille.ColAligns[1] := taRightJustify;
        Grille.ColAligns[2] := taRightJustify;
        Grille.ColAligns[3] := taRightJustify;
        Grille.ColAligns[4] := taRightJustify;
        Grille.RowCount := TobEmploiAug.Detail.Count + 1;
        TobEmploiAug.PutGridDetail(Grille,False,False,'',False);
        Grille.ColWidths[0] := 170;
        Grille.ColWidths[1] := 70;
        Grille.ColWidths[2] := 100;
        Grille.ColWidths[3] := 100;
        Grille.ColWidths[4] := 100;
end ;

Initialization
  registerclasses ( [ TOF_PGAUGMENTATIONCAT ] ) ;
end.

