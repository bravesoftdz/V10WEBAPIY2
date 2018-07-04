{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/06/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFACTIVITEVISAMUL ()
Mots clefs ... : TOF;AFACTIVITEVISAMUL
*****************************************************************}
Unit UTOFAFACTIVITEVISAMUL ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls, HStatus,utob, HQry,HTB97,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche, HDB, mul, DBGrids, db,dbTables,Fe_Main,
{$ENDIF}
     DicoAf,HCtrls,Paramsoc,
     AfUtilArticle,HEnt1, HMsgBox,M3FP, UTOF,UTofAfBaseCodeAffaire ;

Type
  TOF_AFACTIVITEVISAMUL = Class (TOF_AFBASECODEAFFAIRE)   
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;

    private
        MasquerColonnes : Boolean;
        //FPVCol : integer;
        TWC : String;

    public
        GcsTypeVisa:string;

        procedure OnArgument (S : String ) ; override ;
        procedure OnClose                  ; override ;
        procedure OnLoad; override ;
        procedure OnUpdate; override;
        procedure ViserActivite;
        procedure UpdateColumns; // Masquage automatique des colonnes Viseur et Date visa
        procedure SetVISA;
        procedure SetAllVISA;
        {$IFDEF EAGLCLIENT}
        procedure RefreshCrossez;
        procedure OnUpdateNext; override;
        {$ELSE}
        procedure OnGetFieldText(Sender: TField; var Text: string; DisplayText: Boolean);
        {$ENDIF}

  end ;
Procedure AFLanceFiche_Mul_Visa_Activ();

Implementation

procedure TOF_AFACTIVITEVISAMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEVISAMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEVISAMUL.OnArgument (S : String ) ;
var     ComboTypeArticle : THMultiValComboBox;
begin
  Inherited ;

THValComboBox(Getcontrol('ACT_ETATVISA')).value := 'ATT';
GcsTypeVisa:='VIS';
  //mcd 05/03/02
ComboTypeArticle:=THMultiValComboBox(GetControl('ACT_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle(true);
if ComboTypeArticle.Text='' then ComboTypeArticle.Text:=PlusTypeArticleText;
if (GetParamSoc('SO_AFTYPESAISIEACT') <> 'FOL') then // mcd 25/02/03
  begin
  SetControlVisible ('ACT_FOLIO',False);
  SetControlVisible ('TACT_FOLIO',False);
  SetControlVisible ('ACT_FOLIO_',False);
  SetControlVisible ('TACT_FOLIO_',False);
  end;
end ;
                                                  
procedure TOF_AFACTIVITEVISAMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEVISAMUL.ViserActivite;
var i : integer;
    DejaVisa : Boolean;
begin
with TFMul(Ecran) do
   begin
   if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
      begin
      if GcsTypeVisa='VIS' then
            PGIBoxAf('Veuillez sélectionner les lignes d''activité à viser', Caption)
      else
            PGIBoxAf('Veuillez sélectionner les lignes d''activité pour lesquelles vous voulez retirer le visa', Caption);
      exit;
      end;

      {$IFDEF EAGLCLIENT}
      if bSelectAll.Down then
         if not Fetchlestous then
           begin
           bSelectAllClick(Nil);
           bSelectAll.Down := False;
           exit;
           end;
      {$ENDIF}

   if FListe.AllSelected then
      BEGIN
      if GcsTypeVisa='VIS' then
            begin if PGIAskAf('Voulez-vous viser toutes les lignes d''activité ?', Caption) <> mrYes then exit end
      else
            if PGIAskAf('Voulez-vous retirer le visa de toutes les lignes d''activité ?', Caption) <> mrYes then exit;

      if Transactions(SetAllVISA,3) <> oeOK then
          if GcsTypeVisa='VIS' then
                PGIBoxAf('Impossible de viser toutes les lignes d''activité', Caption)
          else
                PGIBoxAf('Impossible de retirer le visa de toutes les lignes d''activité', Caption);

      FListe.AllSelected := false;
      TToolBarButton97(GetControl('bSelectAll')).Down:=false;
      END
   ELSE
      BEGIN
      if GcsTypeVisa='VIS' then
            begin if PGIAskAf('Voulez-vous viser les lignes d''activité sélectionnées ?', Caption) <> mrYes then exit; end
      else
            if PGIAskAf('Voulez-vous retirer le visa des lignes d''activité sélectionnées ?', Caption) <> mrYes then exit;

      DejaVisa := False;
      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
         BEGIN
         FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
         Q.TQ.Seek(FListe.Row-1);
{$ELSE}
{$ENDIF}
         if Q.FindField('ACT_ETATVISA').AsString = GcsTypeVisa then // Ne pas viser les lignes déjà visées
            DejaVisa := True
            else
         if Transactions(SetVISA,3) <> oeOK then
            if GcsTypeVisa='VIS' then
                PGIBoxAf('Impossible de viser une ligne d''activité', Caption)
            else
                PGIBoxAf('Impossible de retirer le visa d''une ligne d''activité', Caption);
         MoveCur(False);
         END;
      FListe.ClearSelected;
      FiniMove;
      if DejaVisa then
            if GcsTypeVisa='VIS' then
                PGIInfoAf('Certaines lignes d''activité déjà visées ont été ignorées', Caption)
            else
                PGIInfoAf('Certaines lignes d''activité qui n''avaient déjà pas de visa ont été ignorées', Caption);
      END;
   ChercheClick;
   end;
end;

// Procédure appelée par Transactions pour mettre à Visé le champ ACT_ETATVISA
procedure TOF_AFACTIVITEVISAMUL.SetVISA;
var Table, TypeActivite, Affaire, NumLigne : String;

// Ressource,DateActivite,Folio,TypeArticle,NumLigne : String;
begin
Table := 'ACTIVITE' ;
with TFMul(Ecran) do
     begin
     TypeActivite := Q.FindField('ACT_TYPEACTIVITE').AsString;
     Affaire := Q.FindField('ACT_AFFAIRE').AsString;
//     Ressource := Q.FindField('ACT_RESSOURCE').AsString;
//     DateActivite := USDateTime(Q.FindField('ACT_DATEACTIVITE').AsDateTime);
//     Folio := Q.FindField('ACT_FOLIO').AsString;
//     TypeArticle := Q.FindField('ACT_TYPEARTICLE').AsString;
//     NumLigne := Q.FindField('ACT_NUMLIGNE').AsString;
     NumLigne := Q.FindField('ACT_NUMLIGNEUNIQUE').AsString;
     end;
ExecuteSQL('UPDATE '+Table+' SET ACT_ETATVISA="'+GcsTypeVisa+'", '+
                                'ACT_VISEUR="'+V_PGI.User+'", '+
                                'ACT_DATEVISA="'+USTime(NowH)+'" '+
                             'WHERE '+
                             'ACT_TYPEACTIVITE="'+TypeActivite+'" AND '+
                             'ACT_AFFAIRE="'+Affaire+'" AND '+
(*                             'ACT_RESSOURCE="'+Ressource+'" AND '+
                             'ACT_DATEACTIVITE="'+DateActivite+'" AND '+
                             'ACT_FOLIO='+Folio+' AND '+
                             'ACT_TYPEARTICLE="'+TypeArticle+'" AND '+  // PL le 14/05/03 : changement de clé de la table activité
                             'ACT_NUMLIGNE='+NumLigne+''); *)
                             'ACT_NUMLIGNEUNIQUE='+NumLigne+'');
end;

// Procédure appelée par Transactions pour mettre à Visé tous les champs ACT_ETATVISA
// en une seule requête
procedure TOF_AFACTIVITEVISAMUL.SetAllVISA;
//var Table : String;
//SelectVisa:string;
begin                             
// PL le 23/01/02

// PL le 22/01/02 c'est toujours la table activite
//Table:='ACTIVITE' ;
// Fin PL le 22/01/02

//if GcsTypeVisa='VIS' then SelectVisa:='ATT' else SelectVisa:='VIS';
ExecuteSQL('UPDATE AFACTIVITEAFFAIRE SET ACT_ETATVISA="'+GcsTypeVisa+'", '+
                                'ACT_VISEUR="'+V_PGI.User+'", '+
                                'ACT_DATEVISA="'+USTime(NowH)+'" '+ TWC );

//                             TWC +
//                             ' AND ACT_ETATVISA="'+SelectVisa+'"');
// Fin PL 23/01/02
end;


procedure TOF_AFACTIVITEVISAMUL.OnLoad;
begin
inherited;
UpdateColumns;
TWC := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;

// Mise à jour de la visibilité des colonnes Viseur et Date visa selon le cas
procedure TOF_AFACTIVITEVISAMUL.UpdateColumns;
begin
with TFMul(Ecran) do
   begin
   MasquerColonnes := (THValComboBox(FindComponent('ACT_ETATVISA')).Value = 'ATT');
   TToolBarButton97(FindComponent('BOUVRIR')).Enabled := (THValComboBox(FindComponent('ACT_ETATVISA')).Value <> 'VIS');
   TToolBarButton97(FindComponent('BREFUSER')).Enabled := (THValComboBox(FindComponent('ACT_ETATVISA')).Value <> 'ATT');
    if TToolBarButton97(FindComponent('BOUVRIR')).Enabled then
        GcsTypeVisa:='VIS'
    else
        GcsTypeVisa:='ATT';

//   ChercheClick;
   end;
end;

// Rétablir le masquage des colonnes à chaque updatage sinon elles réapparaissent
procedure TOF_AFACTIVITEVISAMUL.OnUpdate;
var i : integer;
    //St : String ;
{$IFDEF EAGLCLIENT}
   FPVCol : integer;
{$ENDIF}
begin
inherited;
{$IFDEF EAGLCLIENT}
with TFMul(Ecran).FListe do
  for i := 0 to ColCount-1 do
  begin
     if ((AnsiUpperCase(Cells[i, 0]) = 'VISA') or
         (AnsiUpperCase(Cells[i, 0]) = 'VISEUR') or
         (AnsiUpperCase(Cells[i, 0]) = 'DATE VISA'))
        and MasquerColonnes then ColWidths[i] := 0;
     if (AnsiUpperCase(Cells[i, 0]) = 'LIGNE VISéE') then FPVCol := i;
  end;
RefreshCrossez;
{$ELSE}
with TFMul(Ecran).FListe do
  for i := 0 to Columns.Count-1 do
    with Columns[i] do
      begin
      if (AnsiUpperCase(Title.Caption) = 'VISA') or
         (AnsiUpperCase(Title.Caption) = 'VISEUR') or
         (AnsiUpperCase(Title.Caption) = 'DATE VISA') then Visible := not MasquerColonnes;
      if (AnsiUpperCase(Title.Caption) = 'VISA') then
        Field.OnGetText := OnGetFieldText;
      end;
{$ENDIF}
end;

{$IFDEF EAGLCLIENT}
procedure TOF_AFACTIVITEVISAMUL.OnUpdateNext;
begin
inherited;
RefreshCrossez;
end;

procedure TOF_AFACTIVITEVISAMUL.RefreshCrossez;
var i, j : integer;
   FPVCol : integer;
begin
with TFMul(Ecran).FListe do
  for i := 0 to ColCount-1 do
   if (AnsiUpperCase(Cells[i, 0]) = 'VISA') then
    for j := 1 to RowCount-1 do
     if Cells[i, j] = '-' then
       if Cells[FPVCol, j] = 'VIS' then Cells[i, j] := 'X'
                                   else Cells[i, j] := ' ';
end;
{$ELSE}
// Interception de l'affichage de la colonne Visa pour afficher les coches
procedure TOF_AFACTIVITEVISAMUL.OnGetFieldText(Sender: TField; var Text: string; DisplayText: Boolean);
begin
if TFMul(Ecran).Q.FindField('ACT_ETATVISA').AsString = 'VIS' then Text := 'X' else Text := '-';
end;
{$ENDIF}

procedure AGLViserActivite( Parms : array of variant ; nb : integer );
var  F : TForm;
     LaTof : TOF;
BEGIN
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then LaTof:=TFMul(F).LaTOF else exit;
if (LaTof is TOF_AFACTIVITEVISAMUL) then TOF_AFACTIVITEVISAMUL(LaTof).ViserActivite else exit;

END;

 Procedure AFLanceFiche_Mul_Visa_Activ;
begin
AGLLanceFiche ('AFF','AFACTIVITEVISAMUL','','','');
end;

Initialization
  registerclasses ( [ TOF_AFACTIVITEVISAMUL ] ) ;
  RegisterAglProc('ViserActivite',True,1,AGLViserActivite);
end.
