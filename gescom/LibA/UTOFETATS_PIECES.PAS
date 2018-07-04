{***********UNITE*************************************************
Auteur  ...... :  MC DESSEIGNET
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : etats_PIECES ()
Mots clefs ... : TOF;etats_PIECES
*****************************************************************}
Unit UTOFETATS_PIECES ;
                   
Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
        EQRS1,Maineagl, utob, 
{$ELSE}
    db, dbTables,Qrs1,FE_Main,
{$ENDIF}
    HCtrls, HEnt1, HMsgBox, UTOF,UtofBaseEtats,DicoAf, WIndows, UtilGc ;

Type
  TOF_etats_PIECES = Class (TOF_BASE_ETATS)
      procedure OnArgument(stArgument : String ) ; override ;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
    Private
        procedure ReloadComboEtat;
  end ;
Procedure AFLanceFiche_Portefeuille(Argument:string);
Procedure AFLanceFiche_ListeFact;

Implementation

 procedure TOF_etats_PIECES.ReloadComboEtat;
Var Q : TQuery ;
    i,idef : integer;
    Dernier,StSQL : string;
begin
TFQRS1(ECRAN).FEtat.Items.Clear ; TFQRS1(ECRAN).FEtat.Values.Clear ; idef:=-1;
StSQL:='SELECT MO_TYPE, MO_NATURE, MO_CODE, MO_LIBELLE, MO_LANGUE, MO_DEFAUT FROM MODELES WHERE MO_TYPE="'+TFQRS1(ECRAN).FTypeEtat+'" AND MO_NATURE="'+TFQRS1(ECRAN).FNatEtat+'"' ;
if Not TFQRS1(ECRAN).FChoixEtat then StSQL:=StSQL+' AND MO_CODE="'+TFQRS1(ECRAN).FCodeEtat+'"' ;
Q:=OpenSQL(StSQL,TRUE) ;
While Not Q.EOF do
  BEGIN
  i:=TFQRS1(ECRAN).FEtat.Items.Add(Q.FindField('MO_LIBELLE').AsString) ;
  if Q.FindField('MO_DEFAUT').AsString='X' then Idef:=i ;
  TFQRS1(ECRAN).FEtat.Values.Add(Q.FindField('MO_CODE').AsString) ;
  Q.Next
  END ;
Ferme(Q) ;
TFQRS1(ECRAN).FEtat.ItemIndex:=-1;
if TFQRS1(Ecran).FEtat.Value='' then
  begin
  Dernier:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\'+Apalatys+'\'+NomHalley+'\QRS1Def',GetLeNom(TFQRS1(Ecran).FNomFiltre),'') ;
  if Dernier<>'' then TFQRS1(ECRAN).FEtat.ItemIndex:= TFQRS1(ECRAN).FEtat.values.IndexOf(Dernier);
  if TFQRS1(ECRAN).FEtat.ItemIndex=-1 then TFQRS1(ECRAN).FEtat.ItemIndex:=idef;
  end else TFQRS1(ECRAN).FEtat.ItemIndex:= TFQRS1(ECRAN).FEtat.values.IndexOf(TFQRS1(Ecran).FEtat.Value);
TFQRS1(ECRAN).FEtatClick(Nil) ;
end;

procedure TOF_etats_PIECES.OnArgument(stArgument : String );
var Critere, tmp,St_Plus : string;
    Achat,Venteligne : boolean;
    ZOne : ThValCOmboBox;
BEGIN
Inherited;
Achat :=FAlse;
VenteLigne:=False;
tmp := stArgument;
Critere:=(Trim(ReadTokenSt(tmp)));
While (Critere <>'') do
   begin
   if Critere='ACHATS' then Achat:=True ;
   if Critere='VENTELIG' then Venteligne:=True ;
   Critere:=(Trim(ReadTokenSt(tmp)));
   end;
if Achat then begin
        // mcd afin de traiter le cas du portefeuille appeler pour
        // le module achat et modifier ce qu'il faut dans le mul qui sert aussi en vente
        // gm 13/05/02 j'enléve les critéres relatifs au tiers (car pieces ventes et achat
        //             mélangés, le tiers n'a plus de sens
	SetControlVisible('COMPLEMENT',False);
  SetControlVisible('STATCLIENT',False);
  SetControlVisible('RESSCLIENT',False);
  SetControlVisible('XX_RUPTURE6',False);
  SetControlVisible('TXX_RUPTURE6',False);

  TFQRS1(Ecran).CodeEtat :='APA';
  TFQRS1(Ecran).NatureEtat :='APA';
  Ecran.Caption :=TraduitGa('Etats Ventes/achats par affaire');
  SetControlProperty('GP_NATUREPIECEG','DataType','GCNATUREPIECEG');
  St_Plus:='AND (GPP_NATUREPIECEG="AVC" OR GPP_NATUREPIECEG="FAC"  OR GPP_NATUREPIECEG="APR" OR GPP_NATUREPIECEG="FPR")';
  St_PLus:= St_Plus +  'or (gpp_venteachat="ACH" and gpp_masquernature<>"X")' + AfPlusNatureAchat;
  SetControlProperty('GP_NATUREPIECEG','Plus',St_Plus);
  SetControlProperty('GP_NATUREPIECEG','Text','FAC;FF;AVC');
  SetControlProperty('XX_RUPTURE7','Value','GL_AFFAIRE');
  SetControlProperty('XX_RUPTURE8','Value',' ');
  SetControlProperty('XX_RUPTURE10','Value',' ');
  SetControlProperty('XX_ORDERBY','Value','GL_DATEPIECE');
  (* mcd 04/04/2003 ??? on peut vouloir faire une sélection totale
  Zone := THVALCOmboBOx (GetControl('XX_RUPTURE1'));
  ST_Plus:=ZOne.plus;
  St_Plus:= St_Plus + ' and CO_CODE<>"G02"';
  ZOne.plus:=St_Plus;
  Zone := THVALCOmboBOx (GetControl('XX_RUPTURE2'));
  ST_Plus:=ZOne.plus;
  St_Plus:= St_Plus + ' and CO_CODE<>"G02"';
  ZOne.plus:=St_Plus;  *)
  UpdateCaption(Ecran);
  end
else if Venteligne then begin
        // mcd afin de traiter le cas du portefeuille vente fait à partir des lignes
  Zone:=ThValCOmboBox(GetControl('XX_RUPTURE6'));
  Zone.values[4]:='';
  Zone.items[4]:='';
  TFQRS1(Ecran).CodeEtat :='APV';
  TFQRS1(Ecran).NatureEtat :='APV';
  Ecran.Caption :=TraduitGa('Portefeuille sur ligne');
  SetControlProperty('GP_NATUREPIECEG','DataType','GCNATUREPIECEG');
  SetControlProperty('GP_NATUREPIECEG','Text','FAC;AVC');  // mcd 10/12/02 pour init
  SetControlProperty('XX_RUPTURE10','Value','GL_AFFAIRE');
  SetControlProperty('XX_RUPTURE7','Value','GL_SOUCHE');
  SetControlProperty('XX_RUPTURE8','Value','GL_NUMERO');
  SetControlProperty('XX_ORDERBY','Value',' ');
  Zone := THVALCOmboBOx (GetControl('XX_RUPTURE1'));
  ST_Plus:=ZOne.plus;
  St_Plus:= St_Plus + ' and CO_CODE<>"G02"';
  ZOne.plus:=St_Plus;
  Zone := THVALCOmboBOx (GetControl('XX_RUPTURE2'));
  ST_Plus:=ZOne.plus;
  St_Plus:= St_Plus + ' and CO_CODE<>"G02"';
  ZOne.plus:=St_Plus;
  UpdateCaption(Ecran);
  end
  else   SetControlProperty('GP_NATUREPIECEG','Text','FAC;AVC');  // mcd 10/12/02 pour init

ReloadComboEtat;
END;

procedure TOF_etats_PIECES.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Inherited ;
Aff0:= THEdit(GetControl('GP_AFFAIRE0'));
Aff1:=THEdit(GetControl('GP_AFFAIRE1'));
Aff2:=THEdit(GetControl('GP_AFFAIRE2'));
Aff3:=THEdit(GetControl('GP_AFFAIRE3'));
Aff4:=THEdit(GetControl('GP_AVENANT'));
Tiers:=THEdit(GetControl('GP_Tiers'));
Aff0_:= THEdit(GetControl('GP_AFFAIRE0'));
Aff1_:=THEdit(GetControl('GP_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('GP_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('GP_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('GP_AVENANT_'));
Tiers_:=THEdit(GetControl('GP_Tiers_'));
end ;

Procedure AFLanceFiche_Portefeuille(Argument:string);
begin
AGLLanceFiche ('AFF','AFPORTPIECE','','',Argument);
end;
Procedure AFLanceFiche_ListeFact;
begin
AGLLanceFiche ('AFF','AFLISTEFACT','','','');
end;


Initialization
  registerclasses ( [ TOF_etats_PIECES ] ) ;
end.
