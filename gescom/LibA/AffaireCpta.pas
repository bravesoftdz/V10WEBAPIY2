unit AffaireCpta;


interface

Uses UTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,AffaireUtil,TiersUtil,SysUtils,HEnt1,UtilPGI,Ent1,EntGC,uEntCommun;

Type T_TypeStructAna = (tCodeSection,tLibSection,tLibSsSection );

Procedure GetVentilAffaire (TOBPiece,TOBL,TOBA,TOBTiers,TOBC : TOB; NatV, sRang : String);
Function ChargeAnaAffaire (Axe,Etab : String; TypeChargement: T_TypeStructAna): TOB;
Function TraiteAnaAffaire (Axe : String; TobStruc,TOBPiece,TOBL,TOBA,TOBTiers:TOB; TypeChargement:T_TypeStructAna) : String;
Function RecupChampsTob (Champs : String; TOBPiece,TOBL,TOBA,TOBTiers : TOB) : String;
Function CreatSectionGenere(Cpt,CodeAxe,Etab : string; TOBPiece,TOBL,TOBA,TOBTiers,TobStrucCode : TOB) : Boolean;
Function CreerSousSectionVolee (NumAxe : integer; SSection,Code, Libelle : String) : Boolean;
Function RecupNumPlanSSection (NumAxe,Posit,Lng : integer) : String;
procedure CreatLesSousSections (NumAxe : Integer; Cpt: string ; TobStrucCode,TOBPiece,TOBL,TOBA,TOBTiers : TOB);

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 29/05/2000
Modifié le ... : 29/05/2000
Description .. : Constitution de l'analytique affaire d'une pièce en fonction
Suite ........ : du paramétrage spécifique de ventilation affaire
Mots clefs ... : ANALYTIQUEAFFAIRE
*****************************************************************}
Procedure GetVentilAffaire (TOBPiece,TOBL,TOBA,TOBTiers,TOBC : TOB ;NatV, sRang : String);
Var TOBStruc,TobCpta,TOBAna,TOBEcr : TOB;
    Etab, cpt,  CodeAxe : String;
    NumAxe,MaxAxe,i : integer;
    MajVentil : Boolean;
BEGIN
Cpt := '';
TobStruc := Nil;
MajVentil := False;
if VH_GC.MGCTOBAna = Nil then Exit;
if VH_GC.MGCTOBAna.Detail.count = 0 then Exit;

Etab:=TobPiece.GetValue('GP_ETABLISSEMENT');
if VH_GC.GCVENTAXE2 then MaxAxe := 2 else MaxAxe := 1;

TobCpta := Tob.Create ('Compta gene',Nil,-1);
TOBEcr:=TOBC.Detail[0] ; TobCpta.Dupliquer(TOBEcr,False,True,True);

for NumAxe := 1 to MaxAxe do
    BEGIN
    //TobAna := Nil;
    if TobStruc <> nil then BEGIN TobStruc.Free; TobStruc := Nil; END;
    // Définition du code analytique  (par Etab + Axe puis par Axe seul)
    Codeaxe := 'A' + IntToStr (NumAxe);
    if (Etab <> '')    then TobStruc := ChargeAnaAffaire (CodeAxe,Etab,tCodesection);
    if (TobStruc =Nil) then TobStruc := ChargeAnaAffaire (CodeAxe,'',tCodesection);

    if TobStruc <> Nil then Cpt := TraiteAnaAffaire (CodeAxe,TobStruc,TOBPiece,TOBL,TOBA,TOBTiers,tCodeSection);
    cpt:=copy(cpt,1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
    // Création de la section analytique si nécessaire
    if Not(CreatSectionGenere(Cpt,CodeAxe,Etab,TOBPiece,TOBL,TOBA,TOBTiers,TobStruc)) then Exit ;
    // Création de la ventilation analytique
    if Cpt <> '' then     
       BEGIN
       TobAna := Tob.Create ('VENTIL',TobCpta,-1);
       TobAna.PutValue ('V_SECTION',Cpt); TobAna.PutValue ('V_COMPTE',sRang);
       TobAna.PutValue ('V_NATURE',NatV+IntToStr(NumAxe)); TobAna.PutValue ('V_TAUXMONTANT',100.0);
       TobAna.PutValue ('V_NUMEROVENTIL',1);
       majVentil := True;
       END;
    END; // For sur numAxe

if MajVentil then
   BEGIN // Champs des fils de TOBC / TOBCPTA
   TOBEcr.ClearDetail ;
   for i:=TOBCpta.Detail.Count-1 downto 0 do TOBCpta.Detail[i].ChangeParent(TOBEcr,-1) ;
   END;
// Voir si on peut conserver une ventil d'origine sur 1 axe + ventil affaire sur un second ...
TobStruc.Free; 
TobCpta.Free;
End;

Function CreatSectionGenere(Cpt,CodeAxe,Etab : string; TOBPiece,TOBL,TOBA,TOBTiers,TobStrucCode : TOB) : Boolean;
Var TobStrucLib : Tob;
    Libelle : String;
    NumA : integer;
BEGIN
Result := True; libelle := '';
TobStruclib :=Nil;
if CodeAxe ='A1' then NumA := 1 else NumA := 2;
if Cpt ='' then BEGIN Exit; Result := False; END;
if ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="'+ Cpt + '"') then exit;
// *** Création de la section ***
// Définition du libellé analytique (par Etab + Axe puis par Axe seul)
if Etab <> '' then TobStrucLib := ChargeAnaAffaire (CodeAxe,Etab,tLibSection);
if TobStrucLib= Nil then TobStrucLib := ChargeAnaAffaire(CodeAxe,'',tLibSection);
// Modif BTP
if TobStrucLib= Nil then BEGIN Result := false;exit; END;
Libelle := TraiteAnaAffaire(CodeAxe,TobStrucLib,TOBPiece,TOBL,TOBA,TOBTiers,tLibSection);
if (Libelle = '') then Libelle := Cpt;

Result := CreerSectionVolee ( Cpt,Libelle,NumA );
if Result then
    BEGIN
    // Création des sous sections associées
    CreatLesSousSections (NumA,Cpt,TobStrucCode,TOBPiece,TOBL,TOBA,TOBTiers );
    END;
TobStrucLib.Free;
END;


Function ChargeAnaAffaire (Axe,Etab : String; TypeChargement: T_TypeStructAna): TOB;
Var {stSQL, }TypeStruct : String;
    TobAna,TobDet,TobDetAna : TOB;
    //Q : TQuery;
BEGIN
if TypeChargement = tCodeSection  then TypeStruct := 'SEC' else
if TypeChargement = tLibSection   then TypeStruct := 'LSE' else
if Typechargement = tLibSsSection then TypeStruct := 'LSS';

TobDet:= VH_GC.MGCTOBAna.FindFirst(['AST_AXE','AST_ETABLISSEMENT','AST_TYPESTRUCRANA'],[Axe,Etab,Typestruct],TRUE) ;
if Tobdet <> Nil then TobAna := TOB.Create ('Liste struct ana', Nil,-1) else tobAna := nil;
While TobDet <>Nil do
   BEGIN
   TobDetAna := Tob.Create('STRUCRANAAFFAIRE',TobAna,-1);
   TobDetAna.Dupliquer (TobDet,False,True);
   TobDet:=VH_GC.MGCTOBAna.FindNext(['AST_AXE','AST_ETABLISSEMENT','AST_TYPESTRUCRANA'],[Axe,Etab,Typestruct],TRUE) ;
   END ;
Result := TobAna;
END;

Procedure ChargeTobTiersAff  (CodeTiers : string; TobTiers : Tob);
Var Q : TQuery;
begin
     // fait pour un seul client... on laisse....
Q:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+CodeTiers+'"',True,-1,'',true) ;
if Not Q.EOF then TOBTiers.SelectDB('',Q);
Ferme(Q) ;
end;

Function TraiteAnaAffaire (Axe : String; TobStruc,TOBPiece,TOBL,TOBA,TOBTiers : TOB; TypeChargement: T_TypeStructAna) : String;
Var i, {posit,}Lng, max, NumA : integer;
    TobDet, TobTiersAff : TOB;
    NomTT, Champs, st, Code, CodeAna : String;

BEGIN
//Posit := 0;
CodeAna := '';     Max :=0;
TobTiersAff := Nil;
if Axe ='A1' then NumA := 1 else NumA := 2;
if TypeChargement = tCodeSection  then NomTT := 'AFCHAMPSANA1' else
if TypeChargement = tLibSection   then NomTT := 'AFCHAMPSANALIB1';
if TOBStruc = Nil then Exit;
for i:=0 to TOBStruc.Detail.Count-1 do
    BEGIN
    TOBDet:=TOBStruc.Detail[i];
    Code := TobDet.GetValue('AST_CHAMPS');
    st := RechDom (NomTT,Code,False);
    Champs:=(Trim(ReadTokenSt(st)));
    if st <> '' then Max := StrToInt(Trim(ReadTokenSt(st)));
    if Champs <> '' then
        BEGIN
        if (GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'),'GPP_VENTEACHAT'))='ACH'  then
            begin   // tob tiers de l'affaire à recharger
            if TobTiersAff = Nil then
               begin
               TobTiersAff := Tob.create('TIERS',Nil,-1);
               ChargeTobTiersAff (GetChampsAffaire (TobL.GetValue('GL_AFFAIRE'),'AFF_TIERS'),TobTiersAff);
               end;
            Champs := RecupChampsTob (Champs, TOBPiece,TOBL,TOBA,TOBTiersAff);
            end
        else // tob tiers de la pièce
            Champs := RecupChampsTob (Champs, TOBPiece,TOBL,TOBA,TOBTiers);

        //Posit := strToInt(TobDet.GetValue('AST_POSITION'));
        Lng := StrToInt (TobDet.GetValue('AST_LONGUEUR'));
        if (Max <> 0) And (Max < Lng) then Lng := Max;
        CodeAna := CodeAna + Copy (Champs,1,Lng);
        END;
    END;
if TypeChargement = tCodeSection  then
    Result := BourreLaDonc(CodeAna,TFichierBase(Ord(fbAxe1)+NumA-1))
else Result := CodeAna; //pas de car. de bourrage sur le libellé
if TobTiersAff <> Nil then TobTiersAff.Free;
END;



Function RecupChampsTob (Champs : String; TOBPiece,TOBL,TOBA,TOBTiers : TOB) : String;
Var Prefixe : string;
    Posit : integer;
BEGIN
Posit := Pos('_',Champs);
Prefixe  := copy (Champs,1,Posit-1);
if Prefixe ='T' then // Champs Tiers
    Result := TobTiers.GetValue(Champs) else
if Prefixe = 'YTC' then // Champs complément tiers ressource
    Result := GetChampsComplementTiers(TobTiers.GetValue('T_TIERS'),Champs) else
if Prefixe = 'AFF' then // Champs Affaire
    Result := GetChampsAffaire (TobL.GetValue('GL_AFFAIRE'),Champs) else
if Prefixe = 'GA' then // Champs Article
    Result := TobA.GetValue(Champs) else
if Prefixe = 'GL' then // Champs ligne
    Result := TobL.GetValue(Champs) else
// Modif BTP
if Prefixe = 'BLO' then // Champs ligne
    Result := TobL.GetValue(Champs) else
// ------
if Prefixe = 'GP' then // Champs pièce
    Result := Tobpiece.GetValue(Champs);
END;


procedure CreatLesSousSections (NumAxe : Integer; Cpt: string ; TobStrucCode,TOBPiece,TOBL,TOBA,TOBTiers : TOB);
Var i, Posit, Lng  : integer;
    TobDet : TOB;
    CodeTT, PlanSS,Code,Libelle : String;
    TraiteSS : Boolean;
BEGIN
Code := ''; Libelle := '';
for i:=0 to TOBStrucCode.Detail.Count-1 do
    BEGIN
    TOBDet:=TOBStrucCode.Detail[i];
    if (TobDet.GetValue('AST_CREATSSSECTION')= 'X') then TraiteSS := True else TraiteSS := False;
    if Not(TraiteSS) then Continue;
    CodeTT := TobDet.GetValue('AST_CHAMPS');
    Posit := strToInt(TobDet.GetValue('AST_POSITION'));
    Lng := StrToInt (TobDet.GetValue('AST_LONGUEUR'));
    PlanSS := RecupNumPlanSSection(NumAxe,Posit,Lng );
    if (PlanSS <> '') then
        BEGIN
        // Recup du libellé et création
        Code := Copy (Cpt,Posit,Lng);
        libelle := '';
        CreerSousSectionVolee (NumAxe ,PlanSS,Code, Libelle);
        END;
    END;

END;


Function CreerSousSectionVolee (NumAxe : integer; SSection,Code, Libelle : String) : Boolean;
Var Ax : String;
    Max,NumSSection : integer;
    TOBSS : TOB;
BEGIN
Result := false;
if (Code = '') Or (SSection ='') then Exit;
if (NumAxe <=0) then Exit;
if Libelle = '' then Libelle := Code;
NumSSection := StrToInt(SSection);
// Vérification du code sous section
Max := VH^.SousPlanAxe[TFichierBase(Ord(fbAxe1)+NumAxe-1),NumSSection].Longueur ;
if (Length(Code) > Max) then Exit;
Code:=BourreLaDonc(Code,TFichierBase(Ord(fbAxe1)+NumAxe-1)) ;
// Test de l'existance
Ax:='A'+IntToStr(NumAxe) ;
if ExisteSQL('SELECT PS_CODE FROM SSSTRUCR WHERE PS_AXE="' + Ax + '" AND PS_SOUSSECTION="'
            + SSection + '" AND PS_CODE="' + Code + '"') then Exit ;

TOBSS:=TOB.Create('SSSTRUCR',Nil,-1) ;
TOBSS.PutValue ('PS_AXE',Ax);
TOBSS.PutValue ('PS_SOUSSECTION', SSection);
TOBSS.PutValue ('PS_CODE',Code);
TOBSS.PutValue ('PS_LIBELLE',Copy(Libelle,1,35));
TOBSS.PutValue ('PS_ABREGE',Copy(Libelle,1,17));
Result:=TOBSS.InsertDB(Nil) ;
TOBSS.Free ;
END;

Function RecupNumPlanSSection(NumAxe,Posit,Lng : integer):String;
Var i : integer;
BEGIN
Result := '';
For i:=1 to MaxSousPlan do
    BEGIN
    if (VH^.SousPlanAxe[TFichierBase(Ord(fbAxe1)+NumAxe-1),i].Longueur = Lng ) And
       (VH^.SousPlanAxe[TFichierBase(Ord(fbAxe1)+NumAxe-1),i].Debut = Posit) then
       BEGIN
       Result := VH^.SousPlanAxe[TFichierBase(Ord(fbAxe1)+NumAxe-1),i].Code ;
       Exit;
       END;
    END;
END;

end.
