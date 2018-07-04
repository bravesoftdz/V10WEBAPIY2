{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ZOOMANAL ()
Mots clefs ... : TOF;ZOOMANAL
*****************************************************************}
Unit ZOOMANAL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Menus,
     SaisODA,
     Mul,
     Ent1 ;

Type
  TOF_ZOOMANAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
   private
    Jal, Cpte, Section, QuelAxe,LExo : string ;
    procedure ExoOnChange(sender: TObject);
    procedure BGLOnClick(sender: TObject);
    procedure BVisuPOnClick(sender: TObject);
  end ;

Implementation

uses
{$IFNDEF CCMP}
     QRGLAna,
{$ENDIF}
     CritEdt,
     UtilEdt;


procedure TOF_ZOOMANAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ZOOMANAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ZOOMANAL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_ZOOMANAL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_ZOOMANAL.ExoOnChange ( sender : TObject ) ;
begin
  ExoToDates(GetControlText('Y_EXERCICE'), GetControl('Y_DATECOMPTABLE'),GetControl('Y_DATECOMPTABLE_')) ;
end ;

procedure TOF_ZOOMANAL.BGLOnClick (sender : TObject ) ;
Var D1,D2 : TdateTime ;
    Crit : TCritEdt ;
begin
  inherited;
{$IFNDEF CCMP}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=StrToDate(GetControlText('Y_DATECOMPTABLE')) ; D2:=StrToDate(GetControlText('Y_DATECOMPTABLE_')) ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neGL ;
Crit.Gl.Axe:=GetControlText('Y_AXE') ;
InitCritEdt(Crit) ;
Crit.Cpt1:=GetControlText('Y_SECTION') ; Crit.Cpt2:=Crit.Cpt1 ;
Crit.SQLPLUS:=' AND '+TFMul(Ecran).Q.CRITERES+' ' ;
GLAnalZoom(Crit) ;
{$ENDIF}
end ;

procedure TOF_ZOOMANAL.BVisuPOnClick (sender : TObject ) ;
begin
  TrouveEtLanceSaisieODA(TFMul(Ecran).Q,taConsult) ;
end ;

procedure TOF_ZOOMANAL.OnArgument (S : String ) ;
begin
  Inherited ;
  // récupération des paramètres
  Jal := ReadTokenPipe(S,';' ) ;
  Cpte := ReadTokenPipe(S,';' ) ;
  Section := ReadTokenPipe(S,';' ) ;
  QuelAxe := ReadTokenPipe(S,';' ) ;
  LExo := ReadTokenPipe(S,';' ) ;

  Ecran.Caption := Ecran.Caption + Section ;

  // initialisation des critères
  if Jal<>'' then SetControlText('Y_JOURNAL',Jal )
             else THValComboBox(GetControl('Y_JOURNAL')).ItemIndex := 0 ;
  SetControlText('Y_GENERAL', Cpte ) ;
  SetControlText('Y_SECTION', Section ) ;
  SetControlText('Y_AXE', QuelAxe ) ;
  SetControlEnabled('Y_AXE', false) ;
  SetControlText('Y_EXERCICE', LExo ) ;
  ExoToDates(LExo, GetControl('Y_DATECOMPTABLE'),GetControl('Y_DATECOMPTABLE_')) ;
  THValComboBox(GetControl('Y_QUALIFPIECE')).ItemIndex := 0 ;
  THValComboBox(GetControl('Y_NATUREPIECE')).ItemIndex := 0 ;
  If QuelAxe <> '' Then
   Case QuelAxe[2] of
    '1' : SetControlProperty('Y_SECTION', 'DATATYPE', 'TZSECTION' ) ;
    '2' : SetControlProperty('Y_SECTION', 'DATATYPE', 'TZSECTION2' ) ;
    '3' : SetControlProperty('Y_SECTION', 'DATATYPE', 'TZSECTION3' ) ;
    '4' : SetControlProperty('Y_SECTION', 'DATATYPE', 'TZSECTION4' ) ;
    '5' : SetControlProperty('Y_SECTION', 'DATATYPE', 'TZSECTION5' ) ;
   end ;

  LibellesTableLibre(TTabSheet(GetControl('PZLIBRE')),'TY_TABLE','Y_TABLE','A') ;

  THValComboBox(GetControl('Y_EXERCICE')).OnChange := ExoOnChange ;
  TPopupMenu(GetControl('MENU')).Items[0].OnClick := BGLOnClick ;
  TPopupMenu(GetControl('MENU')).Items[1].OnClick := BVisuPOnClick ;
  TFMul(Ecran).FListe.OnDblClick := BVisuPOnClick ;

end ;

procedure TOF_ZOOMANAL.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ZOOMANAL ] ) ;
end.
