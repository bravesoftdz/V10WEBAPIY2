{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 10/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ANNUBIS (ANNUBIS)
Mots clefs ... : TOM;ANNUBIS
*****************************************************************}
Unit ANNUBIS_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     HTB97,
     UTob ;

Type
  TOM_ANNUBIS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure UpdateBaseWinner;
    procedure ChangeUsage(Sender: TObject);

private
  sGuidPer_c       : string;
  sDos_Winstall    : string;
  sDos_Winouv      : string;
end ;

Implementation

uses
  ulibwindows,DpJurOutils;

procedure TOM_ANNUBIS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_ANNUBIS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ANNUBIS.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_ANNUBIS.OnAfterUpdateRecord ;
begin
  Inherited ;
  // PGR 03/2007 Traitement Mode Partagé pour dossier Winner
  if (sDos_Winstall <> GetControlText('DOS_WINSTALL')) OR
     (sDos_Winouv <> GetControlText('DOS_WINOUV')) then
    UpdateBaseWinner;
end ;

procedure TOM_ANNUBIS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_ANNUBIS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_ANNUBIS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_ANNUBIS.OnArgument ( S: String ) ;
var
  Q1       : TQuery;
begin
  Inherited ;
  sGuidPer_c := TFFiche(Ecran).FLequel;
  Ecran.Caption := 'Organisation : ' + GetNomPer(sGuidPer_c);

  // --- Gestion de la fiche DP DIVERS_WINNER ---

  if (GetControl('DOS_WINSTALL')<>Nil) and (sGuidPer_c<>'') then
  begin
  if ExisteSQL('SELECT 1 FROM DOSSIER WHERE DOS_GUIDPER="'+sGuidPer_c+'"') then
    begin
      // Recherche base Winner
      SetControlVisible('LABELWINNER', True);
      SetControlVisible('DOS_WINSTALL', True);

      Q1 := OpenSql('SELECT DOS_WINSTALL FROM DOSSIER WHERE DOS_GUIDPER="' + sGuidPer_c + '"', True);
      if not Q1.Eof then
        SetControlText('DOS_WINSTALL', Q1.FindField('DOS_WINSTALL').Asstring);
        sdos_winstall := GetControlText('DOS_WINSTALL');
      Ferme(Q1);

      if (FileExists (GetWindowsTempPath+'\cegideweenvexp.ena')) OR (V_PGI.SAV) then
      begin
        SetControlEnabled('LABELWINNER', True);
        SetControlEnabled('DOS_WINSTALL', True);
        //PGR 03/2007 gestion sortie par la croix
        THValComboBox(GetControl('DOS_WINSTALL')).OnChange := ChangeUsage;
      end;

      // PGR 03/2007 Traitement Mode Partagé pour dossier Winner
      SetControlVisible('LABELUSAGE', True);
      SetControlVisible('DOS_WINOUV', True);

      Q1 := OpenSql('SELECT DOS_WINOUV FROM DOSSIER WHERE DOS_GUIDPER="' + sGuidPer_c + '"', True);
      if not Q1.Eof then
        SetControlText('DOS_WINOUV', Q1.FindField('DOS_WINOUV').Asstring);
        THValComboBox(GetControl('DOS_WINOUV')).OnChange := ChangeUsage;
        sdos_winouv := GetControlText('DOS_WINOUV');
      Ferme(Q1);
    end;
  end;

end ;

procedure TOM_ANNUBIS.OnClose ;
begin
  Inherited ;

  // PGR 03/2007 Déplacé dans OnAfterUpdateRecord
  {if sDos_Winstall <> GetControlText('DOS_WINSTALL') then
    UpdateBaseWinner;}
end ;

procedure TOM_ANNUBIS.OnCancelRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :
Description .. : MAJ Base WINNER dans la table DOSSIER,
Suite ........ : qui n'est pas un champ de la tom
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOM_ANNUBIS.UpdateBaseWinner;
var
  BaseWinner: string;
  Q1: TQuery;
begin

  if (GetControl('DOS_WINSTALL')<>Nil) then
  begin
    if GetControlVisible('DOS_WINSTALL') then
    begin
      BaseWinner := '';
      Q1 := OpenSql('SELECT DOS_WINSTALL FROM DOSSIER WHERE DOS_GUIDPER="' + sGuidPer_c + '"', True);

      if not Q1.Eof then
      begin
        BaseWinner := Q1.FindField('DOS_WINSTALL').Asstring;

        if BaseWinner <> GetControlText('DOS_WINSTALL') then
           ExecuteSql('UPDATE DOSSIER SET DOS_WINSTALL = "'+Uppercase(GetControlText('DOS_WINSTALL'))+'" WHERE DOS_GUIDPER="' + sGuidPer_c +'"' );
      end;

      Ferme(Q1);
    end;
  end;

  // PGR 03/2007 Traitement Mode Partagé pour dossier Winner
  if (GetControl('DOS_WINOUV')<>Nil) then
  begin
    if GetControlVisible('DOS_WINOUV') then
    begin
      BaseWinner := '';
      Q1 := OpenSql('SELECT DOS_WINOUV FROM DOSSIER WHERE DOS_GUIDPER="' + sGuidPer_c + '"', True);

      if not Q1.Eof then
      begin
        BaseWinner := Q1.FindField('DOS_WINOUV').Asstring;

        if BaseWinner <> GetControlText('DOS_WINOUV') then
           ExecuteSql('UPDATE DOSSIER SET DOS_WINOUV = "'+Uppercase(GetControlText('DOS_WINOUV'))+'" WHERE DOS_GUIDPER="' + sGuidPer_c +'"' );
      end;

      Ferme(Q1);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PGR
Créé le ...... : 03/2007
Modifié le ... :   /  /
Description .. : Passe en mode édition sur modif de la zone Base Winner et Usage
Mots clefs ... :
*****************************************************************}
procedure TOM_ANNUBIS.ChangeUsage(Sender: TObject);
begin
  // on ne fait rien pendant le OnLoadRecord  ?
  //if bDuringLoad then exit;
  ModeEdition(DS);
  // modif factice pour avoir le message de validation
  SetField('ANB_CHARLIBRE1',GetField('ANB_CHARLIBRE1'));
end;

Initialization
  registerclasses ( [ TOM_ANNUBIS ] ) ;
end.
