unit UTofGSCC;

interface
uses  M3FP, StdCtrls,Controls,Classes, Graphics, forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,uTofAfBaseCodeAffaire,

{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fe_Main,db,dbTables,MajTable,
{$ENDIF}
      UTOF, AglInit, Agent, Ent1;

Type
     TOF_GSCC = Class (TOF_AFBASECODEAFFAIRE)
       private

       public
       procedure OnArgument (stArgument : String )  ; override ;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
       procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.'
                 );
implementation

/////////////////////////////////////////////////////////////////////////////
procedure TOF_GSCC.OnArgument (stArgument : String ) ;
Var St_Plus : string ;
    CC : THValComboBox ;

begin
    Inherited;
St_Plus:='GPP_VENTEACHAT="VEN"';
if not(ctxScot in V_PGI.PGIContexte) then THValComboBox (Ecran.FindComponent('NATUREPIECEG')).Plus := St_Plus ;
//Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');

THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(0,'Facture clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(0,'FAC');
THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(1,'Avoir clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(1,'ZZ2');
THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(2,'Facture + Avoir clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(2,'ZZ1');

SetControlText('NATUREPIECEG','ZZ1');
SetControlText('REFMARGE','DPR');
SetControlEnabled('REFMARGE',false);
THValComboBox (Ecran.FindComponent('GPDEPOT')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GPETABLISSEMENT')).ItemIndex := 0;

if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
      begin SetControlVisible ('TGP_AFFAIRE', False); SetControlVisible ('TGP_AFFAIRE_', False); end;

CC:=THValComboBox(GetControl('GPDOMAINE')) ;
if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GPETABLISSEMENT')) ;
if CC<>Nil then PositionneEtabUser(CC) ;
end;

Procedure TOF_GSCC.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff   := THEdit(GetControl('GPAFFAIRE'))   ; Aff0  := Nil ;
Aff1  := THEdit(GetControl('GPAFFAIRE1'))  ; Aff2  := THEdit(GetControl('GPAFFAIRE2')) ;
Aff3  := THEdit(GetControl('GPAFFAIRE3'))  ; Aff4  := THEdit(GetControl('GPAVENANT'))  ;
Aff_  := THEdit(GetControl('GPAFFAIRE_'))  ; Aff0_ := Nil ;
Aff1_ := THEdit(GetControl('GPAFFAIRE1_')) ; Aff2_ := THEdit(GetControl('GPAFFAIRE2_')) ;
Aff3_ := THEdit(GetControl('GPAFFAIRE3_')) ; Aff4_ := THEdit(GetControl('GPAVENANT_'))  ;
END ;

//////////////////////////////////////////////////////////////////////////////////

procedure TOF_GSCC.OnLoad;
var stWhere, stWhere2, FamilleNiveau1, StFam : string;
    ckb : Tcheckboxstate;
begin
    Inherited;
if GetControlText('NATUREPIECEG') = '' then
    BEGIN
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    END;

stWhere := '';
if GetControlText ('NATUREPIECEG') <> '' then
    begin
    if GetControlText('NATUREPIECEG')='ZZ1' then
       stWhere := ' AND (GP_NATUREPIECEG="FAC" OR GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
       else if GetControlText('NATUREPIECEG')='ZZ2' then
            stWhere := ' AND (GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
            else stWhere := ' AND GP_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
    end;
SetControlText('XX_VARIABLE7',stWhere);

stWhere := '';
FamilleNiveau1 := '';
// Plus besoin Famille NIVEAU 1 car on passe par tablette !
StFam := GetControlText('FAMILLENIV1');
While StFam<>'' do
 FamilleNiveau1 := FamilleNiveau1 +'"'+ ReadTokenSt( StFam ) + '",' ;

if (FamilleNiveau1<>'') and (FamilleNiveau1<>'"<<Tous>>",') then
 BEGIN
 FamilleNiveau1 := Copy(FamilleNiveau1,1,Length(FamilleNiveau1)-1);

 FamilleNiveau1 := '('+FamilleNiveau1+')';
 stWhere:=stWhere+' AND CC_CODE in '+FamilleNiveau1;
 END;
SetControlText('XX_WHERE',stWhere);
stWhere := '';

stWhere2 := '';
FamilleNiveau1 := '';
StFam := GetControlText('FAMILLENIV2');
While StFam<>'' do
 FamilleNiveau1 := FamilleNiveau1 +'"'+ ReadTokenSt( StFam ) + '",' ;

if (FamilleNiveau1<>'') and (FamilleNiveau1<>'"<<Tous>>",') then
 BEGIN
 FamilleNiveau1 := Copy(FamilleNiveau1,1,Length(FamilleNiveau1)-1);

 FamilleNiveau1 := '('+FamilleNiveau1+')';
 stWhere2 := stWhere2+' AND GL_FAMILLENIV2 in '+FamilleNiveau1;
 END;

FamilleNiveau1 := '';
StFam := GetControlText('FAMILLENIV3');
While StFam<>'' do
 FamilleNiveau1 := FamilleNiveau1 +'"'+ ReadTokenSt( StFam ) + '",' ;

if (FamilleNiveau1<>'') and (FamilleNiveau1<>'"<<Tous>>",') then
 BEGIN
 FamilleNiveau1 := Copy(FamilleNiveau1,1,Length(FamilleNiveau1)-1);

 FamilleNiveau1 := '('+FamilleNiveau1+')';
 stWhere2 := stWhere2+' AND GL_FAMILLENIV3 in '+FamilleNiveau1;
 END;
stWhere := stWhere+stWhere2;

// Il faut compléter le Where pour que les sous-requêtes aient les mêmes critères
if GetControlText('GPDOMAINE')<>''       then stWhere:=stWhere+' AND GP_DOMAINE="'+GetControlText('GPDOMAINE')+'"';
// Pas de Filtre REPRESENTANT car on borne sur celui de la requête principale
//if GetControlText('GP_REPRESENTANT')<>''  then stWhere:=stWhere+' AND GP_REPRESENTANT>="'+GetControlText('GP_REPRESENTANT')+'"';
//if GetControlText('GP_REPRESENTANT_')<>'' then stWhere:=stWhere+' AND GP_REPRESENTANT<="'+GetControlText('GP_REPRESENTANT_')+'"';
if GetControlText('GPETABLISSEMENT')<>'' then stWhere:=stWhere+' AND GP_ETABLISSEMENT="'+GetControlText('GPETABLISSEMENT')+'"';
if GetControlText('GPDEPOT')<>''         then stWhere:=stWhere+' AND GP_DEPOT="'+GetControlText('GPDEPOT')+'"';
if GetControlText('GPTIERS')<>''         then stWhere:=stWhere+' AND GP_TIERS>="'+GetControlText('GPTIERS')+'"';
if GetControlText('GPTIERS_')<>''        then stWhere:=stWhere+' AND GP_TIERS<="'+GetControlText('GPTIERS_')+'"';
if GetControlText('GLCODEARTICLE')<>''   then stWhere:=stWhere+' AND GL_CODEARTICLE>="'+GetControlText('GLCODEARTICLE')+'"';
if GetControlText('GLCODEARTICLE_')<>''  then stWhere:=stWhere+' AND GL_CODEARTICLE<="'+GetControlText('GLCODEARTICLE_')+'"';
if GetControlText('GPLIBRETIERS1')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS1="'+GetControlText('GPLIBRETIERS1')+'"';
if GetControlText('GPLIBRETIERS2')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS2="'+GetControlText('GPLIBRETIERS2')+'"';
if GetControlText('GPLIBRETIERS3')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS3="'+GetControlText('GPLIBRETIERS3')+'"';
if GetControlText('GPLIBRETIERS4')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS4="'+GetControlText('GPLIBRETIERS4')+'"';
if GetControlText('GPLIBRETIERS5')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS5="'+GetControlText('GPLIBRETIERS5')+'"';
if GetControlText('GLLIBREART1')<>''     then stWhere:=stWhere+' AND GL_LIBREART1="'+GetControlText('GLLIBREART1')+'"';
if GetControlText('GLLIBREART2')<>''     then stWhere:=stWhere+' AND GL_LIBREART2="'+GetControlText('GLLIBREART2')+'"';
if GetControlText('GLLIBREART3')<>''     then stWhere:=stWhere+' AND GL_LIBREART3="'+GetControlText('GLLIBREART3')+'"';
if GetControlText('GLLIBREART4')<>''     then stWhere:=stWhere+' AND GL_LIBREART4="'+GetControlText('GLLIBREART4')+'"';
if GetControlText('GLLIBREART5')<>''     then stWhere:=stWhere+' AND GL_LIBREART5="'+GetControlText('GLLIBREART5')+'"';
if GetControlText('GLLIBREART6')<>''     then stWhere:=stWhere+' AND GL_LIBREART6="'+GetControlText('GLLIBREART6')+'"';
if GetControlText('GLLIBREART7')<>''     then stWhere:=stWhere+' AND GL_LIBREART7="'+GetControlText('GLLIBREART7')+'"';
if GetControlText('GLLIBREART8')<>''     then stWhere:=stWhere+' AND GL_LIBREART8="'+GetControlText('GLLIBREART8')+'"';
if GetControlText('GLLIBREART9')<>''     then stWhere:=stWhere+' AND GL_LIBREART9="'+GetControlText('GLLIBREART9')+'"';
if GetControlText('GLLIBREARTA')<>''     then stWhere:=stWhere+' AND GL_LIBREARTA="'+GetControlText('GLLIBREARTA')+'"';

// Pièce Vivante ?
//ckb := TCheckBox(Ecran.FindComponent('GP_VIVANTE')).State ;
//if ckb=cbChecked then
// stWhere:=stWhere+' AND GP_VIVANTE="X"'
//else if ckb=cbUnChecked then
// stWhere:=stWhere+' AND GP_VIVANTE<>"X"';

// Il faut virer le premier AND
stWhere := Copy(stWhere,5,Length(stWhere));

SetControlText('XX_VARIABLE12',stWhere);

SetControlText('XX_VARIABLE11',GetControlText('REFMARGE'));

// On force les dates début/fin dans une variable pour contourner DECLA
// qui renvoi No jour !
SetControlText ('GPDATEPIECE', '"'+UsDateTime(StrToDate(GetControlText('GPDATEPIECE0')))+'"');
SetControlText ('GPDATEPIECE_', '"'+UsDateTime(StrToDate(GetControlText('GPDATEPIECE0_')))+'"');

end;

procedure TOF_GSCC.OnUpdate;
begin
    Inherited;

end;

/////////////////////////////////////////////////////////////////////////////

Initialization
registerclasses([TOF_GSCC]) ;
end.

