unit UTofAfBaseLigne_Mul;

interface
uses
{$IFDEF EAGLCLIENT}
     eMul,  Maineagl,
{$ELSE}
     Mul, db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
{$ENDIF}
   StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Hent1,
   HCtrls,UTOF,
   AfUtilArticle, EntGc,
   Dicobtp,UTofAfBasePiece_Mul,M3fp;
      
Type
     TOF_AFBASELIGNE_MUL = Class (TOF_AFBASEPIECE_MUL)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
        procedure AlimPieceVivante; 
     END ;
Type
     TOF_AFLIGNE_MUL = Class (TOF_AFBASEPIECE_MUL)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
     END ;
Procedure AFLanceFiche_Mul_LignePiece(Argument:string);
 Procedure AFLanceFiche_Mul_RegroupLigne(Range,Argument:string);

implementation

procedure TOF_AFBASELIGNE_MUL.OnArgument(stArgument : String );
var ComboTypeArticle:THMultiValComboBox;
BEGIN
Inherited;
SetControlText('GL_TYPELIGNE','ART');
  //mcd 05/03/02
   ComboTypeArticle:=THMultiValComboBox(GetControl('GL_TYPEARTICLE'));
   ComboTypeArticle.plus:=PlusTypeArticle;
END;

procedure TOF_AFBASELIGNE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Inherited;
Aff:=THEdit(GetControl('GL_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('GL_TIERS'));
end;


procedure TOF_AFBASELIGNE_MUL.AlimPieceVivante;
BEGIN
inherited;

END;

procedure AGLAlimLigneVivante( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AFBASELIGNE_MUL) then TOF_AFBASELIGNE_MUL(MaTOF).AlimPieceVivante else exit;
end;

procedure TOF_AFLIGNE_MUL.OnArgument(stArgument : String );
//mcd 12/02/03 var St_Plus : string;
//var zaff : string;
    //Critere, Champ, valeur, LibMul  : String;
    //x : integer;
    //zdateD,ZdateF,zwhere : String;
BEGIN
Inherited;
SetControlVisible ('PCOMPLEMENT',False);
if (Ecran).name='AFREGRLIGNE_MUL' then
      TFMul(Ecran).FiltreDisabled:=true;   //mcd 13/12/02
(* à mettre si on veut regrouper en une seule recherche vente + achat
 attention dans cas, il ne faut pas que le code client soit passer dans les arguements
     // mcd 10/02/03 pour permettre nature piece achat depuis affaire si gérer
  if VH_GC.GAAchatSeria then
    begin
    SetControlProperty('GL_NATUREPIECEG','DataType','GCNATUREPIECEG');
    St_Plus:='AND (GPP_NATUREPIECEG="AVC" OR GPP_NATUREPIECEG="FAC"  OR GPP_NATUREPIECEG="APR" OR GPP_NATUREPIECEG="FPR")';
    if ctxScot in V_PGI.PGIcontexte then St_PLus := St_plus +' or (GPP_NATUREPIECEG="FF")'
     else St_PLus:= St_Plus +  'or (GPP_VENTEACHAT="ACH")';
    SetControlProperty('GL_NATUREPIECEG','Plus',St_Plus);
    SetControlProperty('GL_NATUREPIECEG','Text','FAC;AVC');
    end; *)
END;

procedure TOF_AFLIGNE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Inherited;
Aff:=THEdit(GetControl('GL_AFFAIRE'));
Aff0:=THEdit(GetControl('GL_AFFAIRE0'));
Aff1:=THEdit(GetControl('GL_AFFAIRE1'));
Aff2:=THEdit(GetControl('GL_AFFAIRE2'));
Aff3:=THEdit(GetControl('GL_AFFAIRE3'));
Aff4:=THEdit(GetControl('GL_AVENANT'));
Tiers:=THEdit(GetControl('GL_TIERS'));
end;

Procedure AFLanceFiche_Mul_LignePiece(Argument:string);
begin
AGLLanceFiche ('AFF','AFLIGNE_MUL','','',Argument);
end;
 Procedure AFLanceFiche_Mul_RegroupLigne(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFREGRLIGNE_MUL',Range,'',Argument);
end;


Initialization
registerclasses([TOF_AFBASELIGNE_MUL]);
registerclasses([TOF_AFLIGNE_MUL]);
RegisterAglProc( 'AlimLigneVivante',True,0,AGLAlimLigneVivante);

end.
