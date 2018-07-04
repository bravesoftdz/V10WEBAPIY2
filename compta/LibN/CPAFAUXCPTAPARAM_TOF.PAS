{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPAFAUXCPTAPARAM ()
Mots clefs ... : TOF;CPAFAUXCPTAPARAM
*****************************************************************}
Unit CPAFAUXCPTAPARAM_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,  // Contrôles Halley
     HEnt1,   // TraduireMemoire
     AGLInit, // TheTob
     HMsgBox, // PGIBox
     Utob,    // TOB
     FE_Main, // AGLLanceFiche
     UTOF,
     HTB97     // TOF
     ;

Procedure TLAuxversTLEcrParam ;

Type
  TOF_CPAFAUXCPTAPARAM = Class (TOF)
  private
    Tbl :THvalCombobox ;
    ZoneEcr : THValComboBox ;
    gbCanClose : Boolean;
    procedure BOuvrirOnClick(Sender: TObject);
    procedure RemplitTableLibre(Sender: TObject);  // OK
    procedure RemplitTableLibreEcriture(Sender: TObject); // OK
    Function  LesZonesSontRenseignees : Boolean ;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;


Implementation

Procedure TLAuxversTLEcrParam ;
BEGIN
AGLLanceFiche ('CP','CPAFAUXCPTAPARAM','','','');
END ;

procedure TOF_CPAFAUXCPTAPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTAPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTAPARAM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTAPARAM.OnLoad ;
begin
  Inherited ;
end ;

// Vérifie si les zones sont renseignées
function TOF_CPAFAUXCPTAPARAM.lesZonesSontRenseignees : boolean ;
begin
Result:=False ;
If tbl.text='' then begin
  PgiBox(TraduireMemoire('La table libre tiers n''est pas renseignée.'),TraduireMemoire(Ecran.Caption)) ;
  Exit ;
end ;
If ZoneEcr.text='' then begin
  PgiBox(TraduireMemoire('La zone d''affectation des écritures n''est pas renseignée.'),Traduirememoire(Ecran.Caption));
  Exit ;
end ;
Result:=True ;
end ;

procedure TOF_CPAFAUXCPTAPARAM.BOuvrirOnClick (Sender : TObject) ;
var
  T : TOB;
begin
  gbCanClose:=True;
  // Vérifie si les zones sont renseignées
  if not lesZonesSontRenseignees then begin gbCanClose := False; Exit; end;
  Inherited ;
  T := TOB.Create ('', nil, -1);
  T.AddChampSupValeur('TBL',Tbl.Value);
  T.AddChampSupValeur('ZONEECR',ZoneEcr.Value);
  T.AddChampSupValeur('ECRANA',TCheckBox(GetControl('ECRANA',True)).Checked);
  T.AddChampSupValeur('ECRASER',TCheckBox(GetControl('ECRASER',True)).Checked);
  TheTOB := T;
  Ecran.Close;
end ;

// OK
procedure TOF_CPAFAUXCPTAPARAM.RemplitTableLibre (Sender : TObject) ;
Var i : Integer ;
    Q: TQuery ;
    TobTL : Tob;
begin
TobTL:=Tob.Create('TL',Nil,-1) ;
Q:=OpenSql('SELECT DISTINCT(COUNT(*)),NT_TYPECPTE,CC_LIBELLE FROM NATCPTE '+
           'LEFT JOIN CHOIXCOD ON CC_CODE= NT_TYPECPTE WHERE NT_TYPECPTE LIKE "T%" '+
           'AND CC_TYPE="NAT" GROUP BY NT_TYPECPTE,CC_LIBELLE',True) ;
TobTL.LoadDetailDB('TL','','',Q,True) ;
Ferme(Q) ;
if TobTL.Detail.Count > 0 then
  begin
  Tbl.Clear ;
  For i:=0 To TobTL.Detail.Count-1 do
    begin
    Tbl.Items.add(TobTL.Detail[i].GetValue('CC_LIBELLE')) ;
    Tbl.Values.Add(TobTL.Detail[i].GetValue('NT_TYPECPTE')) ;
    end ;
  end ;
end ;

// OK
procedure TOF_CPAFAUXCPTAPARAM.RemplitTableLibreEcriture (Sender : TObject) ;
Var i : Integer ;
    Q : TQuery ;
    LaValue : String ;
    TobX: Tob ;
begin
TobX:=Tob.Create('X',Nil,-1) ;
Q:=OpenSql('SELECT CC_LIBELLE, CC_CODE FROM CHOIXCOD '+
           'Where CC_type="NAT" and CC_ABREGE="X" and CC_CODE like "E%"', True) ;
TobX.LoadDetailDB('X','','',Q,True) ;
ferme(Q) ;
if TobX.Detail.Count > 0 then
  begin
  For i:=0 To TobX.Detail.Count-1 do
    begin
    ZoneEcr.Items.add(TobX.Detail[i].GetValue('CC_LIBELLE')) ;
    LaValue:='TABLE'+Copy(TobX.Detail[i].GetValue('CC_CODE'),3,1) ;
    ZoneEcr.Values.Add(LaValue) ;
    end ;
  end ;
  TobX.Free ;
end ;

procedure TOF_CPAFAUXCPTAPARAM.OnArgument (S : String ) ;
begin
  gbCanClose := True;
  Tbl:=THvalCombobox(GetControl('TBL',True));
  ZoneEcr:=THValComboBox(GetControl('ZONEECR',True)) ;
  TToolbarButton97(GetControl('BVALIDER',True)).OnClick:=BOuvrirOnClick ;
  RemplitTableLibre(Nil);
  RemplitTableLibreEcriture(Nil) ;
  Inherited ;
end ;

procedure TOF_CPAFAUXCPTAPARAM.OnClose ;
begin
  if not gbCanClose then begin gbCanClose:=True; LastError:=-1; end;
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPAFAUXCPTAPARAM ] ) ;
end.
