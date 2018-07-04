{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ANNDEDOUBLE ()
Mots clefs ... : TOF;ANNDEDOUBLE
*****************************************************************}
Unit UTOFAnnReformate ;

Interface

Uses
   HCtrls, UTOF, UTOB;

{*****************************************************************
// Variables et constantes
*****************************************************************}
const
   csCarDebut_g = '- _';
   csCarMiste_g = '';
   csCarFin_g = '';
   csCarTous_g = '.';
   csLesSigles_g = 'ASS;DR;EURL;EARL;SA;SARL;SRL;SC;SCI;SCM;SNC;' +
                   'M;MME;MELLE;MLLE;MR;ME;MM;ETS;' +
                   'A;à;AU;AUX;L'';LE;LA;LES;D'';DU;DE;DES';

Type
     TOF_ANNREFORMATE = Class (TOF)

       procedure OnArgument ( sParam_p : String ) ; override ;
       procedure OnClose ; override ;

       private
         gdGrille_f : THGrid;
         OBAnnuaire_f : TOB;
         // Index de colonnes
         nColCode_f : integer;
         nColNomA_f : integer;
         nColNomN_f : integer;
         bEnregOk_f : boolean;

         procedure ChargeAnnuaire;
         procedure ValideAnnuaire;

         procedure OnClick_Valider( Sender: TObject );
         procedure OnCellExit_ChangeNom(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean );
end ;


Implementation

uses
{$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ELSE}
{$ENDIF}

   StdCtrls, Classes, ReformateChamp, hstatus, hmsgbox, sysutils, Controls;

{*****************************************************************
Auteur ....... : BM
Procedure  ... : OnArgument
Description .. : Initialisation de la fiche
Paramètres ... :
*****************************************************************}
procedure TOF_ANNREFORMATE.OnArgument ( sParam_p : String ) ;
begin
  Inherited ;
   gdGrille_f := THGrid(GetControl('GRID'));
   gdGrille_f.OnCellExit := OnCellExit_ChangeNom;
   TButton(GetControl('BVALIDER')).OnClick := OnClick_Valider;

   nColCode_f := 0;
   nColNomA_f := 1;
   nColNomN_f := 2;
   gdGrille_f.ColCount := 3;
   gdGrille_f.ColEditables[ nColCode_f ] := false;
   gdGrille_f.ColEditables[ nColNomA_f ] := false;
   gdGrille_f.ColEditables[ nColNomN_f ] := true;
   gdGrille_f.ColWidths[ nColCode_f ] := -1;
   gdGrille_f.ColWidths[ nColNomA_f ] := Round((gdGrille_f.Width / 2)-(gdGrille_f.Width / 40) );
   gdGrille_f.ColWidths[ nColNomN_f ] := gdGrille_f.ColWidths[ nColNomA_f ];

   InitFormats( csCarDebut_g, csCarMiste_g, csCarFin_g, csCarTous_g, csLesSigles_g );
   ChargeAnnuaire;
end;

{*****************************************************************
Auteur ....... : BM
Procedure  ... : OnClose
Description .. : Fermeture de la fiche : enregistrement des données si modifiées
Paramètres ... :
*****************************************************************}
procedure TOF_ANNREFORMATE.OnClose ;
var sMessage_l : string;
begin
Inherited ;
   if bEnregOk_f then
   begin
      sMessage_l := 'Vous n''avez pas enregistré les modifications.' + #13#10 +
                    'Voulez-vous le faire?.' +#13#10;
      if PGIAsk( sMessage_l, 'Dédoublonnage de l''annuaire') = mrYes then
         ValideAnnuaire;
   end;
   OBAnnuaire_f.Free;
   RazFormats;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/11/02
Procédure .... : OnClick_Valider
Description .. : Click sur bouton "VALIDER"
Paramètres ... :
*****************************************************************}
procedure TOF_ANNREFORMATE.OnClick_Valider( Sender: TObject );
begin
  Inherited ;
  ValideAnnuaire;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/11/02
Procédure .... : OnCellExit_ChangeNom
Description .. : En sortie de cellule, traitement la chaîne et initialise l'enregistrement
Paramètres ... : l'objet
                 L'indice de colonne
                 L'indice de ligne
                 Si changée
*****************************************************************}
procedure TOF_ANNREFORMATE.OnCellExit_ChangeNom( Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean );
begin
   if ( nColNomN_f = ACol ) and
      ( gdGrille_f.Cells[nColNomA_f, ARow] <> gdGrille_f.Cells[nColNomN_f, ARow] ) then
   begin
      gdGrille_f.Cells[nColNomN_f, ARow] := ReformateChaine( gdGrille_f.Cells[nColNomN_f, ARow] );
      if ( gdGrille_f.Cells[nColNomA_f, ARow] <> gdGrille_f.Cells[nColNomN_f, ARow] ) then
         bEnregOk_f := true;
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/11/02
Procédure .... : ChargeAnnuaire
Description .. : Charge les données de l'annuaire et reformate le champ NOMPER
Paramètres ... :
*****************************************************************}
procedure TOF_ANNREFORMATE.ChargeAnnuaire;
var
   QRYAnnuaire_l : TQuery;
   nAnnInd_l : integer;   
begin

   OBAnnuaire_f := TOB.Create('ANNUAIRE', nil, -1);
   QRYAnnuaire_l := OpenSQL('select ANN_GUIDPER, ANN_NOMPER from ANNUAIRE order by ANN_NOMPER', true);
   OBAnnuaire_f.LoadDetailDB('ANNUAIRE', '', '', QRYAnnuaire_l, false );
   Ferme(QRYAnnuaire_l);

   bEnregOk_f := false;
   InitMove(OBAnnuaire_f.Detail.Count, 'Dédoublonnage de l''annuaire');
   gdGrille_f.VidePile(false);
   gdGrille_f.RowCount := gdGrille_f.FixedRows + 1;

   for nAnnInd_l := 0 to OBAnnuaire_f.Detail.Count - 1 do
   begin
      gdGrille_f.Cells[nColCode_f, gdGrille_f.FixedRows + nAnnInd_l] := IntToStr(nAnnInd_l);
      gdGrille_f.Cells[nColNomA_f, gdGrille_f.FixedRows + nAnnInd_l] := OBAnnuaire_f.Detail[nAnnInd_l].GetValue('ANN_NOMPER');
      gdGrille_f.Cells[nColNomN_f, gdGrille_f.FixedRows + nAnnInd_l] := ReformateChaine( OBAnnuaire_f.Detail[nAnnInd_l].GetValue('ANN_NOMPER') );

      if gdGrille_f.Cells[nColNomA_f, gdGrille_f.FixedRows + nAnnInd_l] <> gdGrille_f.Cells[nColNomN_f, gdGrille_f.FixedRows + nAnnInd_l] then
         bEnregOk_f := true;

      gdGrille_f.RowCount := gdGrille_f.RowCount + 1;
      MoveCur(False);
   end;
   gdGrille_f.RowCount := gdGrille_f.RowCount - 2 + gdGrille_f.FixedRows;
   FiniMove;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/11/02
Procédure .... : ValideAnnuaire
Description .. : Enregistre les modification dans la base
Paramètres ... :
*****************************************************************}
procedure TOF_ANNREFORMATE.ValideAnnuaire;
var
   nAnnInd_l, nTobInd_l, nAnnTraite_l : integer;
   sOldNomPer_l, sNewNomPer_l, sMessage_l : string;
begin
   InitMove(gdGrille_f.RowCount - gdGrille_f.FixedRows, 'Dédoublonnage de l''annuaire');
   nAnnTraite_l := 0;
   for nAnnInd_l := gdGrille_f.FixedRows to gdGrille_f.RowCount - 1 do
   begin
      nTobInd_l := StrToInt(gdGrille_f.Cells[nColCode_f, nAnnInd_l]);
      sOldNomPer_l := OBAnnuaire_f.Detail[nTobInd_l].GetValue('ANN_NOMPER');
      sNewNomPer_l := gdGrille_f.Cells[nColNomN_f, nAnnInd_l];

      if sOldNomPer_l <> sNewNomPer_l then
      begin
         OBAnnuaire_f.Detail[nTobInd_l].PutValue('ANN_NOMPER', sNewNomPer_l);
         inc(nAnnTraite_l);
      end;

      MoveCur(False);
   end;

   ObAnnuaire_f.UpdateDB(true);
   AvertirTable('ANNUAIRE');
   FiniMove;

   bEnregOk_f := false;
   sMessage_l := IntToStr(OBAnnuaire_f.Detail.Count) + ' élément(s) de l''annuaire traité(s).' + #13#10 +
                 IntToStr(nAnnTraite_l)  + ' élément(s) de l''annuaire modifié(s).' +#13#10;
   PGIInfo( sMessage_l, 'Dédoublonnage de l''annuaire');
   ChargeAnnuaire;
end;


Initialization
   registerclasses ( [ TOF_ANNREFORMATE ] ) ;
end.

