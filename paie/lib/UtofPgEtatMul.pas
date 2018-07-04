{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 08/01/2004
Modifié le ... :   /  /    
Description .. : Etats chaînés : lancement à partir d'une fiche multi-critères.
Mots clefs ... : PAIE
*****************************************************************}
unit UtofPgEtatMul;

interface

uses Classes, StdCtrls, Controls, SysUtils,
{$IFDEF EAGLCLIENT}
     uTob, eMul,Maineagl,
{$ELSE}
     dbtables,Mul,Fe_Main,
{$ENDIF}
     Dialogs,
     comctrls,UTof, HCtrls,
     CritEdt,  // TCritEdtChaine
     AGLInit,  // TheData                             ,
     Filtre,   // Filtre
     Extctrls; // TTimer

procedure AGLLanceFichePGEtatMulChaine( vCritEdtChaine : TCritEdtChaine );

type
  TOF_PGEtatMul = class(TOF)
  private
  protected
    FCritEdtchaine : ClassCritEdtchaine;
  public
    FFiltres : TComBoBox;
    Pages : TPageControl;
    FTimer : TTimer;
    bOpenEtat, bPme: Boolean;
    Arg : String;
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnNew ; override ;
    procedure OnLoad ; override ;
    procedure OnClose; override;
    procedure FTimerTimer(Sender: TObject);
  end ;

implementation
uses lookup,P5Def,PgOutils2,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF GCGC}
//     CalcOle,
     {$ENDIF}
{$ENDIF}
     HEnt1, HTB97, LicUtil;

{ TOF_PGEtatMul }

// GCO
////////////////////////////////////////////////////////////////////////////////
procedure AGLLanceFichePGEtatMulChaine( vCritEdtChaine: TCritEdtChaine );
var lCritEdtChaine : ClassCritEdtChaine;
    st : string;
begin
  lCritEdtChaine := ClassCritEdtChaine.Create;
  lCritEdtchaine.CritEdtChaine := vCritEdtChaine;
  TheData := lCritEdtChaine;
  st := 'CHAINES';
  vCritEdtChaine.UtiliseCritStd := True; // on force les critères
  if vCritEdtChaine.UtiliseCritStd then
    st := st +
          ';'+
          DateToStr(vCritEdtChaine.PGDateDeb)+
          ';'+
          DateToStr(vCritEdtChaine.PGDatefin);
  if vCritEdtChaine.AuFormatPDF then
    st := st +
          ';'+
          'true'
  else
    st := st +
          ';'+
          'false';

//  V_PGI.NoPrintDialog := True; // pas d'affichage de la fenêtre d'impression

  AGLLanceFiche('PAY', vCritEdtChaine.NomFiche, '', '', st);
 end;
////////////////////////////////////////////////////////////////////////////////

procedure TOF_PGEtatMul.OnArgument(Arguments : string) ;

begin
inherited ;
Arg:=Arguments;
if Pos('CHAINES',Arg)<1 then Exit;
// Récupération du TheData pour avoir le CritEdtchaine
FCritEdtchaine := ClassCritEdtchaine(TheData);

if FCritEdtChaine <> nil then
begin
  if FCritEdtchaine.CritEdtchaine.AuFormatPDF then
  begin
    if FCritEdtchaine.CritEdtChaine.MultiPdf then
      V_PGI.QRPDFQueue := ExtractFilePath(FCritEdtchaine.CritEdtChaine.NomPDF) + '\' + PgRendNoDossier() + '-' + ExtractFileName(FCritEdtchaine.CritEdtChaine.NomPDF)
    else
      V_PGI.QRPDFQueue := FCritEdtchaine.CritEdtChaine.NomPDF;

    V_PGI.QRPDFMerge := '' ;
    if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
      V_PGI.QRPDFMerge := V_PGI.QRPDFQueue ;
  end;
end;

FFiltres := TComboBox(GetControl('FFILTRES'));
Pages    := TPageControl(GetControl('PAGES'));

FTimer := TTimer.Create ( Ecran );
FTimer.Enabled := False;
FTimer.OnTimer  := FTimerTimer;
FTimer.Interval := 1000;

{$IFDEF EAGLCLIENT}
{$ELSE}
if not (Ecran is TFMul) then Exit ;
//@@ TFQRS1(Ecran).ChoixEtat:=FALSE ;
{$ENDIF}
bPme:=(not (ctxPCL in V_PGI.PGIContexte)) ;
bOpenEtat:=(V_PGI.PassWord=CryptageSt(DayPass(Date))) and (bPme) ;

OnLoad;
end ;

procedure TOF_PGEtatMul.OnNew;
// var Etat: THValComboBox ;
begin
inherited ;
if Pos('CHAINES',Arg)<1 then Exit;
{$IFNDEF EAGLCLIENT}
if not (Ecran is TFMul) then Exit ;
{$ENDIF}
end;



{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEtatMul.OnLoad;
begin
  inherited;
  if Pos('CHAINES',Arg)<1 then Exit;
  if FCritEdtChaine = nil then Exit;
//  TFQRS1(Ecran).FCodeEtat:=FCritEdtchaine.CritEdtChaine.CodeEtat;
//  TFQRS1(Ecran).FEtat.Value:=FCritEdtchaine.CritEdtChaine.CodeEtat;
  if FCritEdtchaine.CritEdtChaine.Utiliser then
    begin
    // Chargement du filtre du CritEdtChaine
    if FCritEdtChaine.CritEdtChaine.FiltreUtilise<>'' then
       Begin
       ChargeFiltre(FCritEdtchaine.CritEdtChaine.NomFiltre, FFiltres, Pages, True);
       FFiltres.ItemIndex := FFiltres.Items.Indexof( FCritEdtChaine.CritEdtChaine.FiltreUtilise );
       LoadFiltre(FCritEdtchaine.CritEdtChaine.NomFiltre, FFiltres, Pages);
       End;
    FTimer.Enabled := True;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... :   /  /
Description .. : Declencle l' impression de l' état
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEtatMul.FTimerTimer(Sender: TObject);
begin
  FTimer.Enabled := False;

  // On force le Print ou le Preview dans les états chainés
{  if FCritEdtchaine.CritEdtChaine.AuFormatPDF then
    TCheckBox(GetControl('FApercu')).Checked := True
  else
    TCheckBox(GetControl('FApercu')).Checked := False;}

  V_PGI.DefaultDocCopies := FCritEdtchaine.CritEdtChaine.NombreExemplaire;

  TToolBarButton97(GetControl('BOUVRIR')).Click;
  Ecran.Close;

end;

procedure TOF_PGEtatMul.OnClose;
begin
  inherited;
  if Pos('CHAINES',Arg)<1 then Exit;
  TheData := nil;
  FCritEdtChaine.Free;
  FCritEdtChaine := nil;
  FTimer.Free;
  FTimer := nil;
  V_PGI.DefaultDocCopies := 1;
end;

initialization
RegisterClasses([TOF_PGEtatMul]) ;


end.
 