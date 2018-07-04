{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFVISAPIECPRO_MUL ()
Mots clefs ... : TOF;AFVISAPIECPRO_MUL
*****************************************************************}
Unit UTOFAFVISAPIECPRO_MUL ;

Interface

Uses  StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Ent1, HStatus,UTOF,
      utob,HQry,HTB97,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche, HDB, mul, DBGrids, db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      HCtrls, HEnt1, HMsgBox, AffaireUtil,Dicobtp,SaisUtil,M3FP,UTofAfBasePiece_Mul ;

procedure AGLViserFacture( Parms : array of variant ; nb : integer );

Type
  TOF_AFVISAPIECPRO_MUL = Class (TOF_AFBASEPIECE_MUL)
    private
        MasquerColonnes : Boolean;

        TWC : String;
{$IFDEF EAGLCLIENT}
        FPVCol : integer;
{$ENDIF}

    public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnClose                  ; override ;
        procedure OnLoad; override ;
        procedure OnUpdate; override;
        procedure ViserFacture;
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
Procedure AFLanceFiche_Mul_VisaPiecePro(Argument:string);

Implementation

procedure TOF_AFVISAPIECPRO_MUL.OnArgument (stArgument : String ) ;
Var CC : THValComboBox ;
begin
  Inherited ;
THValComboBox(Ecran.FindComponent('GP_ETATVISA')).Value := 'ATT';

CC:=THValComboBox(GetControl('GP_DOMAINE')) ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
end ;

procedure TOF_AFVISAPIECPRO_MUL.OnClose ;
begin
  Inherited ;
end ;


procedure TOF_AFVISAPIECPRO_MUL.ViserFacture;
var i : integer;
    DejaVisa : Boolean;
begin

with TFMul(Ecran) do
   begin
   if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
      begin
      PGIBoxAf('Veuillez sélectionner les pièces à viser', Caption);
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
      if PGIAsk('Voulez-vous viser toutes les pièces ?', Caption) <> mrYes then exit;

      if Transactions(SetAllVISA,3) <> oeOK then PGIBoxAf('Impossible de viser toutes les pièces', Caption);

      FListe.AllSelected := false;
      TToolBarButton97(GetControl('bSelectAll')).Down:=false;
      END
   ELSE
      BEGIN
      if PGIAsk('Voulez-vous viser les pièces sélectionnées ?', Caption) <> mrYes then exit;

      DejaVisa := False;
      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
         BEGIN
         FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
         Q.TQ.Seek(FListe.Row-1);
{$ELSE}
{$ENDIF}
         if Q.FindField('GP_ETATVISA').AsString = 'VIS' then // Ne pas viser les pièces déjà visées
            DejaVisa := True
            else
         if Transactions(SetVISA,3) <> oeOK then PGIBoxAf('Impossible de viser la pièce n° '+Q.FindField('GP_NUMERO').AsString, Caption);
         MoveCur(False);
         END;
      FListe.ClearSelected;
      FiniMove;
      if DejaVisa then PGIInfo('Certaines pièces déjà visées ont été ignorées', Caption);
      END;
   ChercheClick;
   end;
end;

// Procédure appelée par Transactions pour mettre à Visé le champ GP_ETATVISA
procedure TOF_AFVISAPIECPRO_MUL.SetVISA;
var Table,Nature,DateP,Souche,Nb,Indice : String;
begin
Table := 'PIECE' ;
with TFMul(Ecran) do
     begin
     Nature := Q.FindField('GP_NATUREPIECEG').AsString;
     DateP := USDateTime(Q.FindField('GP_DATEPIECE').AsDateTime);
     Souche := Q.FindField('GP_SOUCHE').AsString;
     Nb := Q.FindField('GP_NUMERO').AsString;
     Indice := Q.FindField('GP_INDICEG').AsString;
     end;
ExecuteSQL('UPDATE '+Table+' SET GP_ETATVISA="VIS", '+
                                'GP_VISEUR="'+V_PGI.User+'", '+
                                'GP_DATEVISA="'+USDateTime(NowH)+'" '+
                             'WHERE '+
                             'GP_NATUREPIECEG="'+Nature+'" AND '+
                             'GP_DATEPIECE="'+DateP+'" AND '+
                             'GP_SOUCHE="'+Souche+'" AND '+
                             'GP_NUMERO='+Nb+' AND '+
                             'GP_INDICEG='+Indice+'');
end;

// Procédure appelée par Transactions pour mettre à Visé tous les champs GP_ETATVISA
// en une seule requête
procedure TOF_AFVISAPIECPRO_MUL.SetAllVISA;
var Table : String;
begin
// PL le 22/01/02 c'est toujours la vue AFPIECEAFFAIRE (modifié par PL le 16/10/02)
Table:='AFPIECEAFFAIRE' ;
// Fin PL le 22/01/02
ExecuteSQL('UPDATE '+Table+' SET GP_ETATVISA="VIS", '+
                                'GP_VISEUR="'+V_PGI.User+'", '+
                                'GP_DATEVISA="'+USDateTime(NowH)+'" '+
                             TWC +
                             ' AND GP_ETATVISA="ATT"');
end;

procedure TOF_AFVISAPIECPRO_MUL.OnLoad;
begin
inherited;
UpdateColumns;
TWC := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;

// Mise à jour de la visibilité des colonnes Viseur et Date visa selon le cas
procedure TOF_AFVISAPIECPRO_MUL.UpdateColumns;
begin
with TFMul(Ecran) do
   begin
   MasquerColonnes := (THValComboBox(FindComponent('GP_ETATVISA')).Value = 'ATT');
   TToolBarButton97(FindComponent('BOUVRIR')).Enabled := (THValComboBox(FindComponent('GP_ETATVISA')).Value <> 'VIS');
//   ChercheClick;
   end;
end;

// Rétablir le masquage des colonnes à chaque updatage sinon elles réapparaissent
procedure TOF_AFVISAPIECPRO_MUL.OnUpdate;
var i : integer;
    //St : String ;
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
     if (AnsiUpperCase(Cells[i, 0]) = 'PIèCE VISéE') then FPVCol := i;
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
procedure TOF_AFVISAPIECPRO_MUL.OnUpdateNext;
begin
inherited;
RefreshCrossez;
end;

procedure TOF_AFVISAPIECPRO_MUL.RefreshCrossez;
var i, j : integer;
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
procedure TOF_AFVISAPIECPRO_MUL.OnGetFieldText(Sender: TField; var Text: string; DisplayText: Boolean);
begin
if TFMul(Ecran).Q.FindField('GP_ETATVISA').AsString = 'VIS' then Text := 'X' else Text := '-';
end;
{$ENDIF}

procedure AGLViserFacture( Parms : array of variant ; nb : integer );
var  F : TForm;
     LaTof : TOF;
BEGIN
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then LaTof:=TFMul(F).LaTOF else exit;
if (LaTof is TOF_AFVISAPIECPRO_MUL) then TOF_AFVISAPIECPRO_MUL(LaTof).ViserFacture else exit;

END;

Procedure AFLanceFiche_Mul_VisaPiecePro(Argument:string);
begin
{$IFDEF BTP}
AGLLanceFiche ('AFF','AFVISAPIECPRO_MUL','AFF_AFFAIRE0=I','',Argument);
{$ELSE}
AGLLanceFiche ('AFF','AFVISAPIECPRO_MUL','','',Argument);
{$ENDIF}
end;

Initialization
  registerclasses ( [ TOF_AFVISAPIECPRO_MUL ] ) ;
  RegisterAglProc('ViserFacture',True,1,AGLViserFacture);
end.
