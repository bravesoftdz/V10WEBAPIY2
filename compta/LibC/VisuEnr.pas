unit VisuEnr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Buttons, ExtCtrls, hmsgbox, HSysMenu, HEnt1;

Type TStLue = Record
{ SAARI / HALLEY jusqu'à 120 }
              NumLg,Jal,Date,TP,General,TC,AuxSect,Reference,Libelle,MP,Echeance,S,Montant,TE,NumP : String ;
{ HALLEY au dela de 120 }
              NumEche,Valid,RefExt,DateRefExt,DateCreat,DateModif,Soc,Etab,Affaire,DebE,
              CreE,TauxE,Dev,DebD,CreD,TauxD,DateTauxD,Quot,EcrAN,Qte1,Qte2,Qual1,Qual2,
              RefLibre,TvaEnc,Regime,Tva,TPF,CtrGen,CtrAux,Couv,Let,LetD,DateMin,
              DateMax,RefP,DateP,LPLCR,DateRel,Controle,TotEnc,RelEnc,DateV,RIB,RefRel,
              CouvD,EtatL,NumPInt,Encais,TAN,Eche,Ana,MPaie,NatP : String ;
{ CEGID NORMAL}
              MontantD,MontantE,CodeMontant,Axe : String ;
{ CEGID ETENDU}
              Immo,LT0,LT1,LT2,LT3,LT4,LT5,LT6,LT7,LT8,LT9,TA0,TA1,TA2,TA3,LM0,LM1,LM2,LM3,LD,LB0,LB1,
              Conso,CouvE,LetE : String ;
              End ;

Function VoirEnr (Var Enr : TStLue ; Action : TActionFiche) : Boolean ;

type
  TFVisuEnreg = class(TForm)
    GB1: TGroupBox;
    Pbouton: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    HJal: THLabel;
    HDateC: THLabel;
    HTP: THLabel;
    HGen: THLabel;
    HTC: THLabel;
    HAuxSect: THLabel;
    HReference: THLabel;
    HLibelle: THLabel;
    HMP: THLabel;
    HDateE: THLabel;
    HMontant: THLabel;
    HTM: THLabel;
    HNumP: THLabel;
    HSens: THLabel;
    Jal: TEdit;
    DateC: TEdit;
    TP: TEdit;
    Gen: TEdit;
    TC: TEdit;
    AuxSect: TEdit;
    Reference: TEdit;
    Libelle: TEdit;
    MP: TEdit;
    DateE: TEdit;
    Sens: TEdit;
    Montant: TEdit;
    TM: TEdit;
    NumP: TEdit;
    HMTrad: THSystemMenu;
    HMess: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    Enr : TStLue ;
    ModifFaite : Boolean ;
    Action : TactionFiche ;
    procedure AlimFiche ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Function VoirEnr (Var Enr : TStLue ; Action : TActionFiche) : Boolean ;
Var X : TFVisuEnreg ;
BEGIN
X:=TFVisuEnreg.Create(Application) ;
 Try
  X.Enr:=Enr ;
  X.ModifFaite:=FALSE ;
  X.Action:=Action ;
  X.ShowModal ;
 Finally
  Result:=X.ModifFaite ;
  X.Free ;
 End ;
END ;

procedure TFVisuEnreg.AlimFiche ;
Var  i : Integer ;
BEGIN
Jal.Text:=Enr.Jal; DateC.Text:=Enr.Date; TP.Text:=Enr.TP; Gen.Text:=Enr.general;
TC.Text:=Enr.TC; AuxSect.Text:=Enr.AuxSect; Reference.Text:=Enr.Reference;
Libelle.Text:=Enr.Libelle; MP.Text:=Enr.MP; DateE.Text:=Enr.Echeance; Sens.Text:=Enr.S;
Montant.Text:=Enr.Montant; TM.Text:=Enr.TE ; NumP.Text:=Enr.NumP;
Jal.MaxLength:=Length(Enr.Jal) ; DateC.MaxLength:=Length(Enr.Date) ;
TP.MaxLength:=Length(Enr.TP) ; Gen.MaxLength:=Length(Enr.General) ;
TC.MaxLength:=Length(Enr.TC) ; AuxSect.MaxLength:=Length(Enr.AuxSect) ;
Reference.MaxLength:=Length(Enr.Reference) ; Libelle.MaxLength:=Length(Enr.Libelle) ;
MP.MaxLength:=Length(Enr.MP) ; DateE.MaxLength:=Length(Enr.Echeance) ;
Sens.MaxLength:=Length(Enr.S) ; Montant.MaxLength:=Length(Enr.Montant) ;
TM.MaxLength:=Length(Enr.TE) ; NumP.MaxLength:=Length(Enr.NumP) ;
If Enr.TC='A' Then HAuxSect.Caption:=HMess.mess[2] ;
If Action=taConsult Then
   BEGIN
   For i:=0 To GB1.ControlCount-1 Do If GB1.Controls[i] is TEdit then GB1.Controls[i].Enabled:=FALSE ;
   END ;
END ;

procedure TFVisuEnreg.FormShow(Sender: TObject);
begin
AlimFiche ;
Caption:=Caption+Enr.NumLg ;
UpdateCaption(Self) ;
end;

procedure TFVisuEnreg.BValiderClick(Sender: TObject);
begin
If Action=taConsult Then Exit ;
If HMess.Execute(0,'','')<>mrYes Then Exit ;
ModifFaite:=(Enr.Jal<>Jal.Text) Or (Enr.Date<>DateC.Text) Or (Enr.TP<>TP.Text ) Or
            (Enr.general<>Gen.Text) Or (Enr.TC<>TC.Text) Or (Enr.AuxSect<>AuxSect.Text) Or
            (Enr.Reference<>Reference.Text) Or (Enr.Libelle<>Libelle.Text) Or (Enr.MP<>MP.Text) Or
            (Enr.Echeance<>DateE.Text) Or (Enr.S<>Sens.Text) Or (Enr.Montant<>Montant.Text) Or
            (Enr.TE<>TM.Text) Or (Enr.NumP<>NumP.Text) ;
Enr.Jal:=Jal.Text ; Enr.Date:=DateC.Text ; Enr.TP:=TP.Text ;
Enr.general:=Gen.Text ; Enr.TC:=TC.Text ; Enr.AuxSect:=AuxSect.Text ;
Enr.Reference:=Reference.Text ; Enr.Libelle:=Libelle.Text ; Enr.MP:=MP.Text ;
Enr.Echeance:=DateE.Text ; Enr.S:=Sens.Text ; Enr.Montant:=Montant.Text ;
Enr.TE:=TM.Text ; Enr.NumP:=NumP.Text ;
end;

end.
