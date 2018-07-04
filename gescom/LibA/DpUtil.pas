unit DpUtil;

interface
Uses  AGLInit, 
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fe_Main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF DP}
        Annoutils,
{$ENDIF}
      Forms,HCtrls, SysUtils
    ,Graphics,Controls, stdctrls, spin,UTOM;

    // fonction pour lien DP
//Procedure RechDP(TOM_TIERS : TOM;ecran : TForm; Var CodePer : Integer; Tiers, NatAuxi : String) ;
Procedure RechDP( Var CodePer : Integer; NatAuxi : String; TomT : TOM) ;
Procedure DpSynchro(Maj : Boolean; Var  Codeper : integer; CodeTiers , NatAuxi: String;TomT:TOM) ;
Procedure DPGrise(ecran : TForm) ;
Procedure DPInit(ecran : TForm) ;
Procedure GrisechampEdit( champ: THEdit) ;
Procedure GrisechampSpinEdit( champ: TspinEdit) ;
Procedure GrisechampCombo( champ: TComboBox) ;


implementation


Procedure RechDP( Var CodePer : Integer; NatAuxi:string;TomT : TOM) ;
Var StCodePer : string;
begin
If (NatAuxi <> 'CLI') and (NatAuxi <> 'FOU')then exit;
    // on interdit de ne pas sélectionner de code personne
    // En cas de création, obligation de le faire dans le mul annuaire
//MCD 07/08/00 A revoir quand Setfield sera public. Il faudra faire appel à la fct synchro avec
// Le code TOM
repeat
 //   StCodePer:=AglLanceFiche('YY', 'ANNUAIRE_SEL','ANN_NOMPER='+TomT.GetField('T_ABREGE'),'','TIERS');
 StCodePer:=AglLanceFiche('YY', 'ANNUAIRE_SEL','','','TIERS');
    until StCodePer <> '';

if StCodePer <>'' then
    begin
    CodePer :=StrToInt(StCodePer);
    // a remettre SynchroniseTiers (True, Codeper, TomT.Getfield('T_TIERS'),TomT);
    end;
end;

Procedure DpSynchro(Maj : Boolean; Var  Codeper : integer; CodeTiers , NAtAuxi: String;TomT:TOM) ;
begin
 // Maj = True : Maj Dp vers tiers
 // Maj = False, MAJ Tiers vers DP
If (NatAuxi <> 'CLI') and (NatAuxi <> 'FOU')then exit;
if (CodeTiers = '') then exit;     // cas ou appel en création
//A remettre SynchroniseTiers (Maj, Codeper, Codetiers,TomT);
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
{$IFDEF DP}
SynchroniseTiers (Maj, Codeper, Codetiers);
{$ENDIF}
{$ENDIF}
end;


Procedure DPGrise( ecran : TForm) ;
begin
GriseChampCombo(TComboBox(Ecran.Findcomponent ('T_NATIONALITE')));
GriseChampCombo(TComboBox(Ecran.Findcomponent ('T_LANGUE')));
GriseChampCombo(TComboBox(Ecran.Findcomponent ('T_DEVISE')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_FAX')));
GriseChampSpinEdit( TSpinEdit(Ecran.Findcomponent ('T_MOISCLOTURE')));
GriseChampEdit(THEdit(Ecran.Findcomponent ('T_ABREGE')));
GriseChampEdit(THEdit(Ecran.Findcomponent ('T_JURIDIQUE')));

GriseChampEdit( THEdit(Ecran.Findcomponent ('T_RVA')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_TELEPHONE')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_TELEX')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_APE')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_SIRET')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_LIBELLE')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_PRENOM')));
GriseChampEdit(THEdit(Ecran.Findcomponent ('T_ADRESSE1')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_ADRESSE2')));
GriseChampEdit( THEdit(Ecran.Findcomponent ('T_ADRESSE3')));
GriseChampEdit(THEdit(Ecran.Findcomponent ('T_CODEPOSTAL')));
GriseChampEdit(THEdit(Ecran.Findcomponent ('T_VILLE')));
GriseChampCombo(TComboBox(Ecran.Findcomponent ('T_PAYS')));
end;

Procedure GrisechampEdit(champ: THEdit) ;
begin
    // fct qui gris pour CHamp ThEdit
if (champ <> NIL) then
  begin
  champ.tabstop:=False;
  Champ.Enabled :=False;
  Champ.ReadOnly :=True;
  champ.color :=ClBtnFace;
  end;
end;

Procedure GrisechampSpinEdit(champ: TSpinEdit) ;
begin
    // fct qui gris pour CHamp ThSpinEdit
if (champ <> NIL) then
  begin
  champ.tabstop:=False;
  Champ.Enabled :=False;
  Champ.ReadOnly :=True;
  champ.color :=ClBtnFace;
  end;
end;

Procedure GrisechampCombo(champ: TComboBox) ;
begin
    // fct qui gris pour CHamp THCOmboBOx
if (champ <> NIL) then
  begin
  champ.tabstop:=False;
  Champ.Enabled:=False;
  champ.color :=ClBtnFace;
  end;
end;

Procedure DPInit( ecran : TForm) ;
begin
    //Fct qui permet en cas de saisie avec lien DP
    //d'interdir l'accès à certain champ et de faire rech sur table NAF
// mcd 25/01/01 GriseChampEdit( THEdit(Ecran.Findcomponent ('T_SIRET')));
end;

end.
 
