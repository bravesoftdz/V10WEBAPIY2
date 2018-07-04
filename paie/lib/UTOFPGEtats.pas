{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 14/05/2004
Modifié le ... :   /  /
Description .. : Tof mere gérant les états chaînés
Mots clefs ... : PAIE;ETAT
*****************************************************************}
unit UTOFPGEtats;

// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
{
PT1 14/05/2004 V_50 SB Suite modif gestion Filtre AGL, modification chargement filtre
PT2 15/10/2004 V_50 PH Le paramètrage des états libres n'est pas possible en CWAS
PT3 16/08/2005 V_65 SB FQ 11923 Non acces à la fiche QRS1 en editions chainés
PT4 19/01/2006 V_65 SB FQ 11923 Invisibilité des pagescontrol QRS1 en editions chainés
PT5 30/03/2007 V_72 MF Ajout critères de sélection.
PT6 06/06/2007 V_72 MF FQ 14329
}


interface

uses Classes, StdCtrls,  SysUtils,
  {$IFDEF EAGLCLIENT}
  eQRS1, Maineagl,
  {$ELSE}
  QRS1, Fe_Main,Controls,
  {$ENDIF}
  comctrls, UTof, HCtrls,
  //  Ent1,
  CritEdt, // TCritEdtChaine
  AGLInit, // TheData                             ,
  Extctrls; // TTimer


procedure AGLLanceFichePGEtatChaine(vCritEdtChaine: TCritEdtChaine);

type
  TOF_PGEtats = class(TOF)
  private
  protected
    FCritEdtchaine: ClassCritEdtchaine;
  public
    FFiltres: THValComboBox;   { PT1 }
    Pages: TPageControl;
    FTimer: TTimer;
    bOpenEtat, bPme: Boolean;
    Arg: string;
    procedure OnArgument(Arguments: string); override;
    procedure OnNew; override;
    procedure OnLoad; override;
    procedure OnClose; override;
    procedure FTimerTimer(Sender: TObject);
    procedure OnChangeFiltre ( Sender : TObject ); virtual; { PT1 }
  end;

implementation
uses  PgOutils2, 
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF GCGC}
//  CalcOle,
  {$ENDIF}
  {$ENDIF}
  HEnt1, HTB97, LicUtil;

{ TOF_PGEtats }

// GCO
////////////////////////////////////////////////////////////////////////////////
procedure AGLLanceFichePGEtatChaine(vCritEdtChaine: TCritEdtChaine);
var
  lCritEdtChaine: ClassCritEdtChaine;
  st : string;   // PT5

begin
  lCritEdtChaine := ClassCritEdtChaine.Create;
  lCritEdtchaine.CritEdtChaine := vCritEdtChaine;
  TheData := lCritEdtChaine;

{ d PT5
  if vCritEdtChaine.UtiliseCritStd then
  begin}
    st := 'CHAINES';
    st := st +
          ';'+
          DateToStr(vCritEdtChaine.PGDateDeb)+
          ';'+
          DateToStr(vCritEdtChaine.PGDatefin)+
          ';'+
          vCritEdtChaine.PGNatDeb+
          ';'+
          vCritEdtChaine.PGNatFin+
          ';'+
          vCritEdtChaine.PGEtabDeb+
          ';'+
          vCritEdtChaine.PGEtabFin+
          ';'+
          vCritEdtChaine.CodeEtat;


//  AGLLanceFiche('PAY', vCritEdtChaine.NomFiche, '', vCritEdtChaine.CodeEtat, 'CHAINES;' + DateToStr(vCritEdtChaine.PGDateDeb) + ';' + DateToStr(vCritEdtChaine.PGDatefin))
    AGLLanceFiche('PAY', vCritEdtChaine.NomFiche, '',vCritEdtChaine.CodeEtat, st);
      //vCritEdtChaine.NatureEtat );
{ end
  else}
//    AGLLanceFiche('PAY', vCritEdtChaine.NomFiche, '', vCritEdtChaine.CodeEtat, 'CHAINES'); //vCritEdtChaine.NatureEtat );
// f PT5
end;
////////////////////////////////////////////////////////////////////////////////

procedure TOF_PGEtats.OnArgument(Arguments: string);
{$IFNDEF EAGLCLIENT}
var
  BParam: TToolbarButton97;
{$ENDIF}
begin
  inherited;
  {$IFDEF EAGLCLIENT}
//   TFQRS1(Ecran).ParamEtat := FALSE;
  {$ENDIF}
  Arg := Arguments;
  if Pos('CHAINES', Arg) < 1 then Exit;

  TFQRS1(Ecran).Pages.Enabled := False; { PT3 }

  TFQRS1(Ecran).Pages.Visible := False;  { PT4 }

  // Récupération du TheData pour avoir le CritEdtchaine
  FCritEdtchaine := ClassCritEdtchaine(TheData);

  if FCritEdtChaine <> nil then
  begin
    if FCritEdtchaine.CritEdtchaine.AuFormatPDF then
    begin
      if FCritEdtchaine.CritEdtChaine.MultiPdf then
        // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
        V_PGI.QRPDFQueue := ExtractFilePath(FCritEdtchaine.CritEdtChaine.NomPDF) + '\' + PgRendNoDossier() + '-' + ExtractFileName(FCritEdtchaine.CritEdtChaine.NomPDF)
      else
        V_PGI.QRPDFQueue := FCritEdtchaine.CritEdtChaine.NomPDF;

      V_PGI.QRPDFMerge := '';
      if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
        V_PGI.QRPDFMerge := V_PGI.QRPDFQueue;
    end;
    TFQRS1(Ecran).FCodeEtat := FCritEdtchaine.CritEdtChaine.CodeEtat;
    TFQRS1(Ecran).FEtat.Value := FCritEdtchaine.CritEdtChaine.CodeEtat;
  end;
  { DEB PT1 }
  FFiltres :=THValComboBox(GetControl('FFILTRES'));
  FFiltres.OnChange := OnChangeFiltre;
  { FIN PT1 }
  Pages := TPageControl(GetControl('PAGES'));

  FTimer := TTimer.Create(Ecran);
  FTimer.Enabled := False;
  FTimer.OnTimer := FTimerTimer;
  FTimer.Interval := 1000;

  {$IFDEF EAGLCLIENT}
  {$ELSE}
  if not (Ecran is TFQRS1) then Exit;
  TFQRS1(Ecran).ChoixEtat := FALSE;
  {$ENDIF}
  bPme := (not (ctxPCL in V_PGI.PGIContexte));
  bOpenEtat := (V_PGI.PassWord = CryptageSt(DayPass(Date))) and (bPme);
  if bOpenEtat then
  begin
    {$IFNDEF EAGLCLIENT}
    TFQRS1(Ecran).ParamEtat := TRUE;
    // C'est le boulot de l'AGL !
    BParam := TToolbarButton97(GetControl('BParamEtat'));
    if BParam <> nil then BParam.Visible := TFQRS1(Ecran).ParamEtat;
    {$ENDIF}
  end;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  if bPme then TFQRS1(Ecran).ChoixEtat := TRUE;
  {$ENDIF}
  OnLoad;
end;

procedure TOF_PGEtats.OnNew;
{$IFNDEF EAGLCLIENT}
var
  Etat: THValComboBox;
{$ENDIF}  
begin
  inherited;
  if Pos('CHAINES', Arg) < 1 then Exit;
  {$IFNDEF EAGLCLIENT}
  if not (Ecran is TFQRS1) then Exit;
  Etat := THValComboBox(GetControl('FETAT'));
  if (bPme) and (Etat <> nil) then Etat.Enabled := TFQRS1(Ecran).ChoixEtat;
  {$ENDIF}
  { DEB PT1 }
  if FCritEdtChaine = nil then Exit;
  if FCritEdtchaine.CritEdtChaine.Utiliser then
  begin
    // Chargement du filtre du CritEdtChaine
    if FCritEdtChaine.CritEdtChaine.FiltreUtilise <> '' then
    begin
      // Chargement du Filtre utilisé par les états chainés
      FFiltres.ItemIndex := FFiltres.Items.Indexof( FCritEdtChaine.CritEdtChaine.FiltreUtilise );
      FFiltres.OnChange(FFiltres);
// d PT6
// Quand  CritEdtChaine.FiltreUtilise est renseigné il faut récupérer le code état
// de ce filtre
      FCritEdtChaine.CritEdtChaine.CodeEtat := TFQRS1(Ecran).CodeEtat;
// f PT6
    end;
    FTimer.Enabled := True;
  end;
  { FIN PT1 }
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEtats.OnLoad;
begin
  inherited;
  if Pos('CHAINES', Arg) < 1 then Exit;
  if FCritEdtChaine = nil then Exit;
// d PT6
  if (TFQRS1(Ecran).FCodeEtat = '') then
  begin
// f PT6
    TFQRS1(Ecran).FCodeEtat := FCritEdtchaine.CritEdtChaine.CodeEtat;
    TFQRS1(Ecran).FEtat.Value := FCritEdtchaine.CritEdtChaine.CodeEtat;
  end; // PT6

  { PT1 Code déplacé et modifié dans le OnNew }
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /
Description .. : Declencle l' impression de l' état
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEtats.FTimerTimer(Sender: TObject);
begin
  FTimer.Enabled := False;

  // On force le Print ou le Preview dans les états chainés
  if FCritEdtchaine.CritEdtChaine.AuFormatPDF then
    TCheckBox(GetControl('FApercu')).Checked := True
  else
    TCheckBox(GetControl('FApercu')).Checked := False;

  V_PGI.DefaultDocCopies := FCritEdtchaine.CritEdtChaine.NombreExemplaire;

  TToolBarButton97(GetControl('BValider')).Click;
  Ecran.Close;

end;

procedure TOF_PGEtats.OnClose;
begin
  inherited;
  if Pos('CHAINES', Arg) < 1 then Exit;
  TheData := nil;
  FCritEdtChaine.Free;
  FCritEdtChaine := nil;
  FTimer.Free;
  FTimer := nil;
  V_PGI.DefaultDocCopies := 1;
end;

{ DEB PT1 }
procedure TOF_PGEtats.OnChangeFiltre(Sender: TObject);
begin
  if Ecran is TFQRS1 then
    TFQRS1(Ecran).FFiltresChange(Sender);
end;
{ FIN PT1 }

initialization
  RegisterClasses([TOF_PGEtats]);


end.

