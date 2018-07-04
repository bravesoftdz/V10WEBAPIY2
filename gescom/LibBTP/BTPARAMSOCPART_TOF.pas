{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMSOCPART ()
Mots clefs ... : TOF;BTPARAMSOCPART
*****************************************************************}
Unit BTPARAMSOCPART_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     UtilPGI,
     HMsgBox,
     Graphics,
     UtilLine,
     BTPARAMSOCGEN_TOF,
     UTOF ;


Type
  TOF_BTPARAMSOCPART = Class (TOF_BTPARAMSOCGEN)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	SO_GCNUMARTAUTO : THCheckBox;
    GBPREFIXES,GBSUFFIXES : TGroupBox;
  	function VerifInfos : boolean;
    procedure NUMARTAUTOClick (sender : TObject);
    procedure activeGB(Etat : boolean);
  end ;

Implementation

procedure TOF_BTPARAMSOCPART.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCPART.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCPART.OnUpdate ;
begin
  Inherited ;
  if not VerifInfos then
  begin
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
	StockeInfos(Ecran, LaTob);
end ;

procedure TOF_BTPARAMSOCPART.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCPART.OnArgument (S : String ) ;
begin
  Inherited ;
  SO_GCNUMARTAUTO := THCheckbox(GetControl('SO_GCNUMARTAUTO'));
  SO_GCNUMARTAUTO.OnClick := NUMARTAUTOClick;
  GBPREFIXES := TGroupBox(GetControl('GBPREFIXES'));
  GBSUFFIXES := TGroupBox(GetControl('GBSUFFIXES'));
  chargeEcran(ecran, LaTob);
  activeGB(SO_GCNUMARTAUTO.checked);
end ;

procedure TOF_BTPARAMSOCPART.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCPART.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCPART.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTPARAMSOCPART.VerifInfos: boolean;
var ValArt,ValOuv,ValPre,ValPos : string;
begin
	result := false;
  if SO_GCNUMARTAUTO.Checked then
  begin
  	ValArt := TEdit(GetControl('SO_GCPREFIXEART')).Text;
  	ValOuv := TEdit(GetControl('SO_GCPREFIXENOM')).Text;
  	ValPre := TEdit(GetControl('SO_GCPREFIXEPRE')).Text;
  	ValPos  := TEdit(GetControl('SO_BTPREFIXEPOS')).Text;
    if ((valart<>'') and ((ValArt=ValOuv) or (ValArt=ValPre) or (ValArt = ValPos))) or
    	 ((ValOuv<>'') and ((ValOuv=ValPre) or (ValOuv=ValPos))) or ((ValPre<>'') and (ValPre=ValPos)) then
    begin
    	PGIError ('les préfixes doivent être différent',ecran.caption);
      exit;
    end;
  end;
  result := true;
end;

procedure TOF_BTPARAMSOCPART.NUMARTAUTOClick(sender: TObject);
begin
  activeGB(SO_GCNUMARTAUTO.checked);
end;

procedure TOF_BTPARAMSOCPART.activeGB(Etat: boolean);
begin
	GBPREFIXES.Enabled := Etat;
  GBSUFFIXES.Enabled := Etat;
  THCheckBox(GetCOntrol('SO_GCTYPSUFART')).enabled := Etat;
  if Etat then
  begin
  	TEdit(GetCOntrol('SO_GCPREFIXEART')).Color := clWindow;
  	TEdit(GetCOntrol('SO_GCPREFIXENOM')).Color := clWindow;
  	TEdit(GetCOntrol('SO_GCPREFIXEPRE')).Color := clWindow;
  	TEdit(GetCOntrol('SO_BTPREFIXEPOS')).Color := clWindow;
  	THSpinEdit(GetCOntrol('SO_GCLGNUMART')).Color := clWindow;
  	TEdit(GetCOntrol('SO_GCCOMPTEURART')).Color := clWindow;
  end else
  begin
    TEdit(GetCOntrol('SO_GCPREFIXEART')).Color := clInactiveCaptionText;
  	TEdit(GetCOntrol('SO_GCPREFIXENOM')).Color := clInactiveCaptionText;
  	TEdit(GetCOntrol('SO_GCPREFIXEPRE')).Color := clInactiveCaptionText;
  	TEdit(GetCOntrol('SO_BTPREFIXEPOS')).Color := clInactiveCaptionText;
  	THSpinEdit(GetCOntrol('SO_GCLGNUMART')).Color := clInactiveCaptionText;
  	TEdit(GetCOntrol('SO_GCCOMPTEURART')).Color := clInactiveCaptionText;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPARAMSOCPART ] ) ;
end.

