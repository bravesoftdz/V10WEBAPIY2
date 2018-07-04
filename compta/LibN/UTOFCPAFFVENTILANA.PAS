{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CPAFFVENTILANA ()
Mots clefs ... : TOF;CPAFFVENTILANA
*****************************************************************}
Unit UTOFCPAFFVENTILANA ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF,HTB97,  Ventil, HQry, Mul, Hdb, VentType ;

Type
  TOF_CPAFFVENTILANA = Class (TOF)
    Ligne1 : TcheckBox ;
    Laxe : THValComboBox ;
    CCCode : ThEdit ;
    BCherche : TToolbarButton97 ;
    BZoom : TToolbarButton97 ;
    BSupprimer : TToolbarButton97 ;
    BNouveau : TToolbarButton97 ;
    Procedure BNouveauOnClick (Sender : TObject) ;
    Procedure BSupprimerOnClick (Sender : TObject) ;
    Procedure SupprimeVentil ;
    Procedure BZoomOnClick (Sender : TObject );
    Procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
  end ;

Var LeCodeASupprimer : String ;

Implementation


{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Visualisation et/ou modification des ventilations types
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.BZoomOnClick (Sender : tobject) ;
Var F : TFMul ;
    LeCodeARenvoyer : string ;
BEGIN
F:=TFMul(Ecran) ; LeCodeARenvoyer:='' ;
If F<>Nil then
  If (Not F.Q.Eof) Then
    BEGIN
    If F.Q.FindField('CC_CODE')<>NIL Then LeCodeARenvoyer:=F.Q.FindField('CC_CODE').AsString ;
    END ;
If LeCodeARenvoyer<>'' then Paramventil('TY',LeCodeARenvoyer,'12345',taModif,True) ;
BCherche.Click ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 06/09/2001
Modifié le ... : 06/09/2001
Description .. : Creation nouvelle ventilation type
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.BNouveauOnClick (Sender : tobject) ;
BEGIN
ParamVentiltype ;
BCherche.Click ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 07/09/2001
Modifié le ... :   /  /
Description .. : Suppression d'1 ventil type.
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.BSupprimerOnClick (Sender : tobject) ;
Var F : TFMul ;
BEGIN
F:=TFMul(Ecran) ; LeCodeASupprimer:='' ;
If F<>Nil then
  If (Not F.Q.Eof) Then
    BEGIN
    If F.Q.FindField('CC_CODE')<>NIL Then
     BEGIN
     LeCodeASupprimer:=F.Q.FindField('CC_CODE').AsString ;
     END;
    END ;
If PGIAsk(TraduireMemoire('Confirmez-vous la suppression ?'),TraduireMemoire('Suppression Code Ventilation type'))=mrYes then
  BEGIN
  If LeCodeASupprimer<>'' then
  If Transactions(SupprimeVentil,1) = oeUnknown then PGIBox(TraduireMemoire('Une erreur est survenue, traitement interrompu'),TraduireMemoire('Suppression code ventilation type')) ;
  END ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 07/09/2001
Modifié le ... : 07/09/2001
Description .. : Suppression de la ventil type selon son axe
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPAFFVENTILANA.SupprimeVentil  ;
var ReqSql : String ;
BEGIN
Try
  BeginTrans ;
  ReqSql := 'DELETE FROM VENTIL WHERE V_COMPTE IN (SELECT CC_CODE FROM CHOIXCOD ';
  ReqSql := ReqSql +'WHERE CC_TYPE="VTY") AND V_NATURE LIKE "TY%" AND V_COMPTE="'+LeCodeASupprimer+'"';
  ExecuteSql(ReqSql) ;
  ReqSql := 'DELETE FROM CHOIXCOD WHERE CC_CODE ="'+LeCodeASupprimer+'" AND CC_TYPE="VTY"';
  ExecuteSql(ReqSql) ;
  CommitTrans ;
  BCherche.Click ;
Except
  Rollback ;
End ;
END ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : G.Verdier
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. : Recherche des enregistrements selon l'axe choisi.
Mots clefs ... :
*****************************************************************}
procedure TOF_CPAFFVENTILANA.OnLoad ;
var AxeChoisi : Integer ;
    TY , XX : String ;
BEGIN
TY := 'TY%' ;
If LAXE <> NIL then
  BEGIN
  AxeChoisi := THValComboBox(GetControl('LAXE')).ItemIndex ;
  Case AxeChoisi of
    0 : TY :='TY%' ;
    1 : TY :='TY1' ;
    2 : TY :='TY2' ;
    3 : TY :='TY3' ;
    4 : TY :='TY4' ;
    5 : TY :='TY5' ;
    else TY :='TY%' ;
  END ; // End du Case
  TY := 'V_NATURE LIKE "'+TY+'"' ;
  END ; // End du If LAXE
If LIGNE1 <> NIL then
  BEGIN
  if GetControlText('LIGNE1')='X' then XX := 'AND V_NUMEROVENTIL=1' else XX:='' ;
  END ; // End du If LIGNE
SetControlText('XX_WHERE',TY+XX) ;
Inherited ;
END ;

procedure TOF_CPAFFVENTILANA.OnArgument (S : String ) ;
begin
Inherited ;
Laxe := ThValComboBox(GetControl('LAXE')) ;
CCCode := ThEdit(GetControl('CC_CODE')) ;
Laxe.Itemindex := 0 ;
SetFocusControl('CC_CODE') ;
Ligne1 := TCheckBox(GetControl('LIGNE1'));
BCherche := TToolBarButton97(GetControl('BCHERCHE')) ;
Bsupprimer := TToolBarButton97(GetControl('BSUPPRIMER')) ;
BZoom := TToolBarButton97(GetControl('BZOOM')) ;
BNouveau := TToolBarButton97(GetControl('BINSERT')) ;
If BZoom <> NIL then
   BEGIN
   BZoom.OnClick := BZoomOnClick ;
   END ;
If BNouveau <> NIL then
   BEGIN
   BNouveau.OnClick := BNouveauOnClick ;
   END ;
If BSupprimer <> NIL then
   BEGIN
   Bsupprimer.OnClick := BSupprimerOnClick ;
   END ;
end ;


Initialization
  registerclasses ( [ TOF_CPAFFVENTILANA ] ) ;
end.
