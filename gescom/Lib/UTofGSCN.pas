unit UTofGSCN;

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
     TOF_GSCN = Class (TOF_AFBASECODEAFFAIRE)
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
procedure TOF_GSCN.OnArgument (stArgument : String ) ;
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
THValComboBox (Ecran.FindComponent('GP_DEPOT')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GP_ETABLISSEMENT')).ItemIndex := 0;

if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
      begin SetControlVisible ('TGP_AFFAIRE', False); SetControlVisible ('TGP_AFFAIRE_', False); end;

CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
if CC<>Nil then PositionneEtabUser(CC) ;
end;

Procedure TOF_GSCN.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff   := THEdit(GetControl('GP_AFFAIRE'))   ; Aff0  := Nil ;
Aff1  := THEdit(GetControl('GP_AFFAIRE1'))  ; Aff2  := THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3  := THEdit(GetControl('GP_AFFAIRE3'))  ; Aff4  := THEdit(GetControl('GP_AVENANT'))  ;
Aff_  := THEdit(GetControl('GP_AFFAIRE_'))  ; Aff0_ := Nil ;
Aff1_ := THEdit(GetControl('GP_AFFAIRE1_')) ; Aff2_ := THEdit(GetControl('GP_AFFAIRE2_')) ;
Aff3_ := THEdit(GetControl('GP_AFFAIRE3_')) ; Aff4_ := THEdit(GetControl('GP_AVENANT_'))  ;
END ;

//////////////////////////////////////////////////////////////////////////////////

procedure TOF_GSCN.OnLoad;
var stWhere, stWhere2, FamilleNiveau1, StFam : string;
begin
    Inherited;
if GetControlText('NATUREPIECEG') = '' then
    BEGIN
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    END;

///////////////////////////////////////////////////////////////////////////////
// MAJ XX_WHERE pour Filtre auto Decla
///////////////////////////////////////////////////////////////////////////////
stWhere := '';
if GetControlText ('NATUREPIECEG') <> '' then
    begin
    if GetControlText('NATUREPIECEG')='ZZ1' then
       stWhere := ' AND (GP_NATUREPIECEG="FAC" OR GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
       else if GetControlText('NATUREPIECEG')='ZZ2' then
            stWhere := ' AND (GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
            else stWhere := ' AND GP_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
    end;
SetControlText('XX_WHERE',stWhere);

///////////////////////////////////////////////////////////////////////////////
// MAJ XX_VARIABLE12 pour Filtre sur Piece GP_
///////////////////////////////////////////////////////////////////////////////
// Il faut compléter le Where pour que les sous-requêtes aient les mêmes critères
stWhere := '';
if GetControlText('GP_DOMAINE')<>''       then stWhere:=stWhere+' AND GP_DOMAINE="'+GetControlText('GP_DOMAINE')+'"';
if GetControlText('GP_REPRESENTANT')<>''  then stWhere:=stWhere+' AND GP_REPRESENTANT>="'+GetControlText('GP_REPRESENTANT')+'"';
if GetControlText('GP_REPRESENTANT_')<>'' then stWhere:=stWhere+' AND GP_REPRESENTANT<="'+GetControlText('GP_REPRESENTANT_')+'"';
if GetControlText('GP_ETABLISSEMENT')<>'' then stWhere:=stWhere+' AND GP_ETABLISSEMENT="'+GetControlText('GP_ETABLISSEMENT')+'"';
if GetControlText('GP_DEPOT')<>''         then stWhere:=stWhere+' AND GP_DEPOT="'+GetControlText('GP_DEPOT')+'"';
if GetControlText('GP_TIERS')<>''         then stWhere:=stWhere+' AND GP_TIERS>="'+GetControlText('GP_TIERS')+'"';
if GetControlText('GP_TIERS_')<>''        then stWhere:=stWhere+' AND GP_TIERS<="'+GetControlText('GP_TIERS_')+'"';
if GetControlText('GP_LIBRETIERS1')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS1="'+GetControlText('GP_LIBRETIERS1')+'"';
if GetControlText('GP_LIBRETIERS2')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS2="'+GetControlText('GP_LIBRETIERS2')+'"';
if GetControlText('GP_LIBRETIERS3')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS3="'+GetControlText('GP_LIBRETIERS3')+'"';
if GetControlText('GP_LIBRETIERS4')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS4="'+GetControlText('GP_LIBRETIERS4')+'"';
if GetControlText('GP_LIBRETIERS5')<>''   then stWhere:=stWhere+' AND GP_LIBRETIERS5="'+GetControlText('GP_LIBRETIERS5')+'"';
stWhere := '';
FamilleNiveau1 := '';
// Famille NIVEAU 1
StFam := GetControlText('FAMILLENIV1');
While StFam<>'' do
 FamilleNiveau1 := FamilleNiveau1 +'"'+ ReadTokenSt( StFam ) + '",' ;

if (FamilleNiveau1<>'') and (FamilleNiveau1<>'"<<Tous>>",') then
 BEGIN
 FamilleNiveau1 := Copy(FamilleNiveau1,1,Length(FamilleNiveau1)-1);

 FamilleNiveau1 := '('+FamilleNiveau1+')';
 stWhere:=stWhere+' AND CC_CODE in '+FamilleNiveau1;
 END;


// Il faut virer le premier AND
stWhere := Copy(stWhere,5,Length(stWhere));

// Pour requête autonome sur GP_
SetControlText('XX_VARIABLE12',stWhere);

SetControlText('XX_VARIABLE11',GetControlText('REFMARGE'));

///////////////////////////////////////////////////////////////////////////////
// MAJ XX_VARIABLE7 pour Filtre sur Ligne GL_
///////////////////////////////////////////////////////////////////////////////
stWhere := '';
if GetControlText('GLCODEARTICLE')<>''    then stWhere:=stWhere+' AND GL_CODEARTICLE>="'+GetControlText('GLCODEARTICLE')+'"';
if GetControlText('GLCODEARTICLE_')<>''   then stWhere:=stWhere+' AND GL_CODEARTICLE<="'+GetControlText('GLCODEARTICLE_')+'"';
if GetControlText('GLLIBREART1')<>''      then stWhere:=stWhere+' AND GL_LIBREART1="'+GetControlText('GLLIBREART1')+'"';
if GetControlText('GLLIBREART2')<>''      then stWhere:=stWhere+' AND GL_LIBREART2="'+GetControlText('GLLIBREART2')+'"';
if GetControlText('GLLIBREART3')<>''      then stWhere:=stWhere+' AND GL_LIBREART3="'+GetControlText('GLLIBREART3')+'"';
if GetControlText('GLLIBREART4')<>''      then stWhere:=stWhere+' AND GL_LIBREART4="'+GetControlText('GLLIBREART4')+'"';
if GetControlText('GLLIBREART5')<>''      then stWhere:=stWhere+' AND GL_LIBREART5="'+GetControlText('GLLIBREART5')+'"';
if GetControlText('GLLIBREART6')<>''      then stWhere:=stWhere+' AND GL_LIBREART6="'+GetControlText('GLLIBREART6')+'"';
if GetControlText('GLLIBREART7')<>''      then stWhere:=stWhere+' AND GL_LIBREART7="'+GetControlText('GLLIBREART7')+'"';
if GetControlText('GLLIBREART8')<>''      then stWhere:=stWhere+' AND GL_LIBREART8="'+GetControlText('GLLIBREART8')+'"';
if GetControlText('GLLIBREART9')<>''      then stWhere:=stWhere+' AND GL_LIBREART9="'+GetControlText('GLLIBREART9')+'"';
if GetControlText('GLLIBREARTA')<>''      then stWhere:=stWhere+' AND GL_LIBREARTA="'+GetControlText('GLLIBREARTA')+'"';
stWhere2 := '';
FamilleNiveau1 := '';
// Famille NIVEAU 2
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
// Famille NIVEAU 3
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

// Il faut virer le premier AND
stWhere := Copy(stWhere,5,Length(stWhere));

// Pour requête autonome sur GL_
SetControlText('XX_VARIABLE7',stWhere);

// On force les dates début/fin dans une variable pour contourner DECLA
// qui renvoi No jour !
SetControlText ('DATEPIECE', UsDateTime(StrToDate(GetControlText('GPDATEPIECE0'))));
SetControlText ('DATEPIECE_', UsDateTime(StrToDate(GetControlText('GPDATEPIECE0_'))));

// Il nous faut la date de fin demois précédent la période demandée
SetControlText ('DATDEBM1', UsDateTime(StrToDate(GetControlText('GPDATEPIECE0'))-1));

end;

procedure TOF_GSCN.OnUpdate;
begin
    Inherited;

end;

/////////////////////////////////////////////////////////////////////////////

Initialization
registerclasses([TOF_GSCN]) ;
end.

