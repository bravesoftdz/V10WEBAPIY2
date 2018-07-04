unit VentilCpta;


interface

Uses UTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,TiersUtil,
{$IFDEF AFFAIRE}
     AffaireUtil,
{$ENDIF}
     SysUtils,HEnt1,UtilPGI,Ent1,EntGC,ParamSoc,uEntCommun;

Type T_TypeStructAna = (tCodeSection,tLibSection,tLibSsSection );

Procedure GetVentilGescom (TOBPiece,TOBL,TOBA,TOBTiers,TOBC : TOB; NatV, sRang : String;TOBGCS : TOB=nil);
Function ChargeAnaGescom (Axe,Etab,VenteAchat : String; TypeChargement: T_TypeStructAna): TOB;
Function TraiteAnaGescom (Axe : String; TobStruc,TOBPiece,TOBL,TOBA,TOBTiers:TOB; TypeChargement:T_TypeStructAna;TOBGCS : TOB=nil) : String;
Function RecupChampsTob (Champs : String; TOBPiece,TOBL,TOBA,TOBTiers,TOBConso : TOB) : String;
Function RecupLibelle (NomChamps,Champs : String; TOBPiece,TOBL,TOBA,TOBTiers,TOBTab : TOB) : String;
Function CreatSectionGenere(Cpt,CodeAxe,Etab : string; TOBPiece,TOBL,TOBA,TOBTiers,TobStrucCode : TOB) : Boolean;
Function CreerSousSectionVolee (NumAxe : integer; SSection,Code, Libelle : String) : Boolean;
Function RecupNumPlanSSection (NumAxe,Posit,Lng : integer) : String;
Function CreatLesSousSections (NumAxe : Integer; Cpt: string ; TobStrucCode,TOBPiece,TOBL,TOBA,TOBTiers : TOB) : Boolean;
Function BourreChamps (Champs: String; Lng:Integer;LeType : TFichierBase; CarBour:boolean): String ;
Procedure GetVentilPort (TOBPiece,TOBTiers,TOBC : TOB; NatV, sRang : String;TOBGCS : TOB=nil);
Procedure GetVentilAnaOD (TOBPiece,TOBL,TOBA,TOBTiers,TOBC : TOB; NatV, sRang : String;TOBGCS : TOB=nil);
// Analytique sur consoammations
function CreerSectionConso ( Axe,Etab : string; TOBCOnso,TOBAffaire,TOBtiers,TOBArticle,TOBressource: TOB; typeStruc: T_TypeStructAna): string;
Function TraiteAnaConso (Axe : String; TOBStruc,TOBCOnso,TOBAffaire,TOBtiers,TOBArticle,TOBressource : TOB ; TypeChargement:T_TypeStructAna) : String;
Function RecupChampsTobGenerique (Champs : String; TOBPiece,TOBL,TOBAFF,TOBA,TOBTiers,TOBConso : TOB) : String;
// -----------------------
implementation
uses FactComm,BTPUtil,CalcOLEGenericBTP;


Function TraiteAnaConso (Axe : String; TOBStruc,TOBCOnso,TOBAffaire,TOBtiers,TOBArticle,TOBressource : TOB ; TypeChargement:T_TypeStructAna) : String;
Var i, Lng, NumA : integer;
    TobDet, TobTab  : TOB;
    NomTT, Champs, CodeAna,IsLibelle,NomChamps,CodeVide : String;
    QQ: TQuery ;
    CodCha, SvgCodeAff : String;
    Aff0,Aff1,Aff2,Aff3,Aff4 : string;
BEGIN
  CodeAna := '';
  CodeVide := '';
  SvgCodeAff := '';
  CodCha:='';
  // -
  if Axe ='A1' then NumA := 1
  else if Axe ='A2' then NumA := 2
  else NumA :=3;   //mcd 21/11/02 maxi 3 axe géré en GC
  if TOBStruc = Nil then Exit;
  {$IFDEF AFFAIRE}
  If (Copy(TOBConso.Getvalue('BCO_AFFAIRE'),1,1) = 'W') then
  begin
    CodCha:=GetChampsAffaire (TOBConso.Getvalue('BCO_AFFAIRE'), 'AFF_CHANTIER');
    if CodCha <> '' then
    begin
    	SvgCodeAff := TOBConso.Getvalue('BCO_AFFAIRE');
    	TOBConso.SetString('BCO_AFFAIRE',CodCha);
      BTPCodeAffaireDecoupe (TOBConso.GetValue('BCO_AFFAIRE'),Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
      TOBConso.SetString ('BCO_AFFAIRE0',Aff0);
      TOBConso.SetString ('BCO_AFFAIRE1',Aff1);
      TOBConso.SetString ('BCO_AFFAIRE2',Aff2);
      TOBConso.SetString ('BCO_AFFAIRE3',Aff3);
    end;
  end;
  {$ENDIF}
  if TypeChargement = tCodeSection  then NomTT := 'GCCHAMPSANA' else
  if TypeChargement = tLibSection   then NomTT := 'GCCHAMPSANA';
  // --
  TobTab:=Tob.Create('_Tablettes',nil,-1);
  QQ:=OpenSQL('Select CO_LIBELLE,CO_LIBRE from Commun where CO_TYPE="GCS" and CO_LIBRE Like "%;CON;%"',True) ;
  if Not QQ.Eof then TobTab.LoadDetailDB('COMMUN','','',QQ,False,True) ;
  Ferme (QQ) ;
  // ---
  for i:=0 to TobTab.Detail.count-1 do
  begin
    Champs:=TobTab.Detail[i].GetValue('CO_LIBRE');
    TobTab.Detail[i].AddChampSup('_CHAMPS',False);
    TobTab.Detail[i].PutValue('_CHAMPS',ReadTokenSt(Champs)) ;
  end ;
  Champs:='' ;
  for i:=0 to TOBStruc.Detail.Count-1 do
  BEGIN
    TOBDet:=TOBStruc.Detail[i];
    if (TypeChargement=tCodeSection) and (TOBDet.GetValue('GDA_TYPESTRUCTANA')='LSE') then continue ;
    if (TypeChargement=tLibSection) and (TOBDet.GetValue('GDA_TYPESTRUCTANA')='SEC') then continue ;
    Champs:=TOBDet.GetValue('GDA_LIBCHAMP') ;
    NomChamps:=Champs ;
    IsLibelle := TobDet.GetValue('GDA_ISLIBELLE') ;
    if Champs <> '' then
    BEGIN
      Champs := RecupChampsTobGenerique (Champs, nil,nil,TOBAffaire,TOBArticle,TOBTiers,TOBCOnso);
      If IsLibelle='X' then Champs := RecupLibelle (NomChamps,Champs, nil,nil,TOBArticle,TOBTiers,TOBTab);
      Lng := StrToInt (TobDet.GetValue('GDA_LONGUEUR'));
      // on renseigne le CodeVide avec le caractère de bourrage pour comparer
      // avec le code section en fin de fonction
      CodeVide := CodeVide + BourreChamps ('',Lng,TFichierBase(Ord(fbAxe1)+NumA-1),True);
      // Bourrage effectué avec le caractère de bourrage uniquement pour le code
      // pour le libellé on met des espaces
      if Lng <> 0 then
      if TypeChargement = tLibSection then Champs:=BourreChamps (Champs,Lng,TFichierBase(Ord(fbAxe1)+NumA-1),False)
                                      else Champs:=BourreChamps (Champs,Lng,TFichierBase(Ord(fbAxe1)+NumA-1),True);
      CodeAna := CodeAna + Copy (Champs,1,Lng);
    END;
  END;
  if TypeChargement = tCodeSection  then
  Begin
    // Si le code ne contient que des caractères de bourrage, on le vide pour
    // pouvoir utiliser la section par défaut
    if CodeAna = CodeVide then CodeAna := '';

    // Si le code est renseigné, on complète avec le caractère de bourrage
    if length(CodeAna)>0 then Result := BourreLaDonc(CodeAna,TFichierBase(Ord(fbAxe1)+NumA-1))
                         else Result := '';
  End else
  begin
    // cadrage à gauche du libellé
    Result := Trim(CodeAna);
  end;

  // Traitement particuliers pour les appels d'interventions : BRL 041207
  //    Restauration du code affaire dans la tob LIGNE
  If (SvgCodeAff <> '') then
  begin
    Aff0 := ''; Aff1 := ''; Aff2 := ''; Aff3 := ''; Aff4 := '';
  	TobConso.PutValue('BCO_AFFAIRE',SvgCodeAff);
    BTPCodeAffaireDecoupe (TOBConso.GetString('BCO_AFFAIRE'),Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
    TOBConso.SetString ('BCO_AFFAIRE0',Aff0);
    TOBConso.SetString ('BCO_AFFAIRE1',Aff1);
    TOBConso.SetString ('BCO_AFFAIRE2',Aff2);
    TOBConso.SetString ('BCO_AFFAIRE3',Aff3);
  end;

  TobTab.Free;
END;



function CreerSectionConso ( Axe,Etab : string; TOBCOnso,TOBAffaire,TOBtiers,TOBArticle,TOBressource: TOB; typeStruc: T_TypeStructAna): string;
var TobStruc : TOB;
begin
  result := '';
  TobStruc := nil;
  if (Etab <> '') then TobStruc := ChargeAnaGescom(Axe, Etab, 'CON', typeStruc);
  if (TobStruc = nil) then TobStruc := ChargeAnaGescom (Axe, '', 'CON', typeStruc);
  if (TobStruc <> nil) then result := TraiteAnaConso(Axe, TobStruc, TOBCOnso,TOBAffaire,TOBtiers,TOBArticle,TOBressource, typeStruc);
  if TobStruc <> nil then TobStruc.free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 29/05/2000
Modifié le ... : 29/05/2000
Description .. : Constitution de l'analytique affaire d'une pièce en fonction
Suite ........ : du paramétrage spécifique de ventilation affaire
Mots clefs ... : ANALYTIQUEAFFAIRE
*****************************************************************}
Procedure GetVentilGescom (TOBPiece,TOBL,TOBA,TOBTiers,TOBC : TOB; NatV, sRang : String;TOBGCS : TOB=nil);


	function VentilLigneOk(TOBL : TOB; Axe : string) : boolean;
  var TOBAnaP,TOBPreAna : TOB;
  begin
  	result := false;
    if TOBL.detail.count = 0 then exit;
    TOBAnaP := TOBL.detail[0];
    TOBPreAna := TOBANAP.findFirst(['YVA_AXE'],[Axe],true);
    if TOBPreAna = nil then exit;
    if (TOBPreAna.getValue('YVA_SECTION')='') or (TOBPreAna.getValue('YVA_POURCENTAGE')=0) then exit;
    result := true;
  end;

Var TOBStruc,TobCpta,TOBAna,TobEcr  : TOB;
    Etab, cpt,  CodeAxe ,VenteAchat: String;
    NumAxe,MaxAxe,I : integer;
    MajVentil : Boolean;
BEGIN
Cpt := '';
TobStruc := Nil;
MajVentil := False;
if GetParamSoc('SO_GCDESACTIVECOMPTA') then Exit ;
if VH_GC.MGCTOBAna = Nil then Exit;
if VH_GC.MGCTOBAna.Detail.count = 0 then Exit;
  // mcd 13/01/2003 : le source ne traite pas l'analytique Achat !!!
VenteAchat:=GetInfoParPiece(TobPiece.getvalue('GP_NATUREPIECEG'),'GPP_VENTEACHAT') ;

Etab := TobPiece.GetValue('GP_ETABLISSEMENT');
if VH_GC.GCVENTAXE3 then MaxAxe := 3          // mcd 21/11/03 pour gérer les 3 axes
else if VH_GC.GCVENTAXE2 then MaxAxe := 2 else MaxAxe := 1;

TobCpta := Tob.Create ('Compta gene',Nil,-1);
// TobCpta.Dupliquer (TobC,True,True,True);
TOBEcr:=TOBC.Detail[0] ; TobCpta.Dupliquer(TOBEcr,False,True,True);//mcd  21/11/02
for NumAxe := 1 to MaxAxe do
    BEGIN
    cpt:=''; // mcd 21/11/02 pour remise à blanc cpt entre 2 traitement de
    if TobStruc <> nil then BEGIN TobStruc.Free; TobStruc := Nil; END;
    // Définition du code analytique  (par Etab + Axe puis par Axe seul)
    Codeaxe := 'A' + IntToStr (NumAxe);
    // -- MODIF LS Pour recuperer la ventilation mise en place au niveau de la ligne de piece
    if VentilLigneOk(TOBL,CodeAxe) then continue;
    // --
    if (Etab <> '')    then TobStruc := ChargeAnaGescom (CodeAxe,Etab,VenteAchat,tCodesection);
    if (TobStruc =Nil) then TobStruc := ChargeAnaGescom (CodeAxe,'',VenteAchat,tCodesection);

    if TobStruc <> Nil then Cpt := TraiteAnaGescom (CodeAxe,TobStruc,TOBPiece,TOBL,TOBA,TOBTiers,tCodeSection,TOBGCS);
    cpt:=copy(cpt,1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
    // -------------------------------------------------
    // Ajout LS 22/02/2013 Pour gestion analytique stock
    // -------------------------------------------------
    if (cpt='') and
    (GetParamSocSecur('SO_OPTANALSTOCK',false)) and
    (Pos(TOBPiece.GetString('GP_NATUREPIECEG'),'FF;AFS')>0) and
    (TOBl.GetString('GL_TYPEARTICLE')='MAR') and
    (TOBl.GetString('GL_TENUESTOCK')='X') and
    (TOBl.GetInteger('GL_IDENTIFIANTWOL')=0) then
    begin
      cpt:=copy(BTGetSectionStock(TOBL.GetString('GL_DEPOT')),1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
    end;
    // -------------------------------------------------
    // Création de la section analytique si nécessaire
    if Not(CreatSectionGenere(Cpt,CodeAxe,Etab,TOBPiece,TOBL,TOBA,TOBTiers,TobStruc)) then Exit;
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

  if majVentil then
  BEGIN // Champs des fils de TOBC / TOBCPTA
    TOBEcr.ClearDetail ;  //mcd 21/11/02
    for i:=TOBCpta.Detail.Count-1 downto 0 do TOBCpta.Detail[i].ChangeParent(TOBEcr,-1) ;  //mcd 21/11/02
  END;
  // Voir si on peut conserver une ventil d'origine sur 1 axe + ventil affaire sur un second ...
  TobStruc.Free;
  TobCpta.free;
End;

Procedure GetVentilAnaOD (TOBPiece,TOBL,TOBA,TOBTiers,TOBC : TOB; NatV, sRang : String;TOBGCS : TOB=nil);
Var TOBStruc,TobCpta,TOBAna,TobEcr  : TOB;
    Etab, cpt,  CodeAxe ,VenteAchat: String;
    NumAxe,MaxAxe,I : integer;
    MajVentil : Boolean;
BEGIN
Cpt := '';
TobStruc := Nil;
MajVentil := False;
if GetParamSoc('SO_GCDESACTIVECOMPTA') then Exit ;
if VH_GC.MGCTOBAna = Nil then Exit;
if VH_GC.MGCTOBAna.Detail.count = 0 then Exit;
  // mcd 13/01/2003 : le source ne traite pas l'analytique Achat !!!
VenteAchat:=GetInfoParPiece(TobPiece.getvalue('GP_NATUREPIECEG'),'GPP_VENTEACHAT') ;

Etab := TobPiece.GetValue('GP_ETABLISSEMENT');
if VH_GC.GCVENTAXE3 then MaxAxe := 3          // mcd 21/11/03 pour gérer les 3 axes
else if VH_GC.GCVENTAXE2 then MaxAxe := 2 else MaxAxe := 1;

TobCpta := Tob.Create ('Compta gene',Nil,-1);
// TobCpta.Dupliquer (TobC,True,True,True);
TOBEcr:=TOBC.Detail[0] ; TobCpta.Dupliquer(TOBEcr,False,True,True);//mcd  21/11/02
for NumAxe := 1 to MaxAxe do
    BEGIN
    cpt:=''; // mcd 21/11/02 pour remise à blanc cpt entre 2 traitement de
    if TobStruc <> nil then BEGIN TobStruc.Free; TobStruc := Nil; END;
    // Définition du code analytique  (par Etab + Axe puis par Axe seul)
    Codeaxe := 'A' + IntToStr (NumAxe);
    // --
    if (Etab <> '')    then TobStruc := ChargeAnaGescom (CodeAxe,Etab,VenteAchat,tCodesection);
    if (TobStruc =Nil) then TobStruc := ChargeAnaGescom (CodeAxe,'',VenteAchat,tCodesection);

    if TobStruc <> Nil then Cpt := TraiteAnaGescom (CodeAxe,TobStruc,TOBPiece,TOBL,TOBA,TOBTiers,tCodeSection,TOBGCS);
    cpt:=copy(cpt,1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
    // -------------------------------------------------
    // Ajout LS 22/02/2013 Pour gestion analytique stock
    // -------------------------------------------------
    if (cpt='') and
    	 (GetParamSocSecur('SO_OPTANALSTOCK',false)) and
    	 (Pos(TOBPiece.GetString('GP_NATUREPIECEG'),'FF;AFS')>0) and
       (TOBl.GetString('GL_TYPEARTICLE')='MAR') and
       (TOBl.GetString('GL_TENUESTOCK')='X') and
       (TOBl.GetInteger('GL_IDENTIFIANTWOL')=0) then
    begin
    	cpt:=copy(BTGetSectionStock(TOBL.GetString('GL_DEPOT')),1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
    end;
    // -------------------------------------------------
    // Création de la section analytique si nécessaire
    if Not(CreatSectionGenere(Cpt,CodeAxe,Etab,TOBPiece,TOBL,TOBA,TOBTiers,TobStruc)) then Exit;
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

if majVentil then
    BEGIN // Champs des fils de TOBC / TOBCPTA
    {TobaVider := Tob.Create ('ventil a sup',Nil,-1);
    For i := TOBC.Detail.Count-1 downto 0 do
        BEGIN
        TOBDet:=TOBC.Detail[i];
        TOBDet.ChangeParent(TOBaVider,-1);
        END;
    TobaVider.Free; }
//mcd   TOBC.ClearDetail ;
//mcd    TobC.Dupliquer (TobCpta,True,True,True);
    TOBEcr.ClearDetail ;  //mcd 21/11/02
    for i:=TOBCpta.Detail.Count-1 downto 0 do TOBCpta.Detail[i].ChangeParent(TOBEcr,-1) ;  //mcd 21/11/02

    {For i := TOBCpta.Detail.Count-1 downto 0 do
        BEGIN
        TOBDet:=TOBCpta.Detail[i];
        TOBDet.ChangeParent(TOBC,-1);
        END;}
    END;
// Voir si on peut conserver une ventil d'origine sur 1 axe + ventil affaire sur un second ...
TobStruc.Free;
TobCpta.free;
End;

Procedure GetVentilPort (TOBPiece,TOBTiers,TOBC : TOB; NatV, sRang : String;TOBGCS : TOB=nil);
Var TOBStruc,TobCpta,TOBAna,TobEcr  : TOB;
    Etab, cpt,  CodeAxe ,VenteAchat: String;
    NumAxe,MaxAxe,I : integer;
    MajVentil : Boolean;
BEGIN
  Cpt := '';
  TobStruc := Nil;
  MajVentil := False;
  if GetParamSoc('SO_GCDESACTIVECOMPTA') then Exit ;
  if VH_GC.MGCTOBAna = Nil then Exit;
  if VH_GC.MGCTOBAna.Detail.count = 0 then Exit;
  // mcd 13/01/2003 : le source ne traite pas l'analytique Achat !!!
  VenteAchat:=GetInfoParPiece(TobPiece.getvalue('GP_NATUREPIECEG'),'GPP_VENTEACHAT') ;

  Etab := TobPiece.GetValue('GP_ETABLISSEMENT');
  if VH_GC.GCVENTAXE3 then MaxAxe := 3          // mcd 21/11/03 pour gérer les 3 axes
  else if VH_GC.GCVENTAXE2 then MaxAxe := 2 else MaxAxe := 1;

  TobCpta := Tob.Create ('Compta gene',Nil,-1);
  // TobCpta.Dupliquer (TobC,True,True,True);
  TOBEcr:=TOBC.Detail[0] ; TobCpta.Dupliquer(TOBEcr,False,True,True);//mcd  21/11/02
  for NumAxe := 1 to MaxAxe do
  BEGIN
    cpt:=''; // mcd 21/11/02 pour remise à blanc cpt entre 2 traitement de
    if TobStruc <> nil then BEGIN TobStruc.Free; TobStruc := Nil; END;
    // Définition du code analytique  (par Etab + Axe puis par Axe seul)
    Codeaxe := 'A' + IntToStr (NumAxe);
    // --
    if (Etab <> '')    then TobStruc := ChargeAnaGescom (CodeAxe,Etab,VenteAchat,tCodesection);
    if (TobStruc =Nil) then TobStruc := ChargeAnaGescom (CodeAxe,'',VenteAchat,tCodesection);

    if TobStruc <> Nil then Cpt := TraiteAnaGescom (CodeAxe,TobStruc,TOBPiece,nil,nil,TOBTiers,tCodeSection,TOBGCS);
    cpt:=copy(cpt,1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
    // Création de la section analytique si nécessaire
    if Not(CreatSectionGenere(Cpt,CodeAxe,Etab,TOBPiece,nil,nil,TOBTiers,TobStruc)) then Exit;
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

  if majVentil then
  BEGIN // Champs des fils de TOBC / TOBCPTA
    TOBEcr.ClearDetail ;  //mcd 21/11/02
    for i:=TOBCpta.Detail.Count-1 downto 0 do TOBCpta.Detail[i].ChangeParent(TOBEcr,-1) ;  //mcd 21/11/02
  END;
  // Voir si on peut conserver une ventil d'origine sur 1 axe + ventil affaire sur un second ...
  TobStruc.Free;
  TobCpta.free;
End;


Function CreatSectionGenere(Cpt,CodeAxe,Etab : string; TOBPiece,TOBL,TOBA,TOBTiers,TobStrucCode : TOB) : Boolean;
Var TobStrucLib : Tob;
    Libelle ,VenteAchat: String;
    NumA : integer;
BEGIN
Result := True; libelle := '';
if CodeAxe ='A1' then NumA := 1
  else if CodeAxe ='A2' then NumA := 2
  else NumA :=3;   //mcd 21/11/02 maxi 3 axe géré en GC
if Cpt ='' then BEGIN Exit; Result := False; END;
  // mcd 13/01/2003 : le source ne traite pas l'analytique Achat !!!
VenteAchat:=GetInfoParPiece(TobPiece.getvalue('GP_NATUREPIECEG'),'GPP_VENTEACHAT') ;

// mcd 21/11/02 teste en plus axe ... idem fct créersectionvolee if ExisteSQL('Select S_SECTION from SECTION where S_SECTION="'+ Cpt + '"') then exit;
if VH^.LiaisonY2ViaShare then
begin
  if ExisteSQL('Select CSP_SECTION from CSECTION where CSP_SECTION="'+ Cpt +'" AND CSP_AXE="A'+IntToStr(NumA)+'"') then exit;
end else
begin
  if ExisteSQL('Select S_SECTION from SECTION where S_SECTION="'+ Cpt +'" AND S_AXE="A'+IntToStr(NumA)+'"') then exit;

end;
// *** Création de la section ***
// Définition du libellé analytique (par Etab + Axe puis par Axe seul)
if Etab <> '' then
     TobStrucLib := ChargeAnaGescom (CodeAxe, Etab,VenteAchat, tLibSection)
else TobStrucLib := Nil ;
if TobStrucLib = Nil then TobStrucLib := ChargeAnaGescom(CodeAxe, '',VenteAchat, tLibSection) ;
Libelle := TraiteAnaGescom(CodeAxe,TobStrucLib,TOBPiece,TOBL,TOBA,TOBTiers,tLibSection) ;
if Libelle = '' then Libelle := Cpt ;

Result := CreerSectionVolee ( Cpt,Libelle,NumA );
if Result then
    BEGIN
    // Création des sous sections associées
    if not VH^.LiaisonY2ViaShare then CreatLesSousSections (NumA,Cpt,TobStrucCode,TOBPiece,TOBL,TOBA,TOBTiers );
    END;
TobStrucLib.Free;
END;


Function ChargeAnaGescom (Axe,Etab,VenteAchat : String; TypeChargement: T_TypeStructAna): TOB;
Var TypeStruct : String;
    TobAna,TobDet,TobDetAna : TOB;
BEGIN
if TypeChargement = tCodeSection  then TypeStruct := 'SEC' else
if TypeChargement = tLibSection   then TypeStruct := 'LSE' else
if Typechargement = tLibSsSection then TypeStruct := 'LSS';
TobDet:= VH_GC.MGCTOBAna.FindFirst(['GDA_AXE','GDA_ETABLISSEMENT','GDA_TYPESTRUCRANA','GDA_TYPECOMPTE'],[Axe,Etab,Typestruct,VenteAchat],TRUE) ;
if Tobdet <> Nil then TobAna := TOB.Create ('Liste struct ana', Nil,-1) else TobAna := nil;
While TobDet <>Nil do
   BEGIN
   // mcd 21/11/02  ??? STRUCTANA n'existe pas TobDetAna := Tob.Create('STRUCTANA',TobAna,-1);
   TobDetAna := Tob.Create('DECOUPEANA',TobAna,-1);
   TobDetAna.Dupliquer (TobDet,False,True);
   TobDet:=VH_GC.MGCTOBAna.FindNext(['GDA_AXE','GDA_ETABLISSEMENT','GDA_TYPESTRUCRANA','GDA_TYPECOMPTE'],[Axe,Etab,Typestruct,VenteAchat],TRUE) ;
   END ;
Result := TobAna;
END;

Procedure ChargeTobTiersAff  (CodeTiers : string; TobTiers : Tob);
Var Q : TQuery;
begin
Q:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+CodeTiers+'"',True) ;
if Not Q.EOF then TOBTiers.SelectDB('',Q);
Ferme(Q) ;
end;

Function TraiteAnaGescom (Axe : String; TobStruc,TOBPiece,TOBL,TOBA,TOBTiers:TOB; TypeChargement:T_TypeStructAna;TOBGCS : TOB=nil) : String;
Var i, Lng, NumA : integer;
    TobDet, TobTab  : TOB;
    NomTT, Champs, CodeAna,IsLibelle,NomChamps,CodeVide : String;
    QQ: TQuery ;
    CodCha, SvgCodeAff,prefixe : String;
BEGIN
  if TOBL <> nil then prefixe := GetPrefixeTable (TOBL);
CodeAna := '';
CodeVide := '';
if Axe ='A1' then NumA := 1
  else if Axe ='A2' then NumA := 2
  else NumA :=3;   //mcd 21/11/02 maxi 3 axe géré en GC
if TOBStruc = Nil then Exit;

// Traitement particuliers pour les appels d'interventions : BRL 041207
//    Le code affaire (celui de l'appel) est remplacé dans la tob ligne par le code chantier associé
//    si celui-ci est renseigné.
// 		Ceci afin de respecter la structure du code analytique paramétrée.
CodCha:='';
{$IFDEF AFFAIRE}
If (Copy(TOBPiece.Getvalue('GP_AFFAIRE'),1,1) = 'W') then
begin
  CodCha:=GetChampsAffaire (TOBPiece.Getvalue('GP_AFFAIRE'), 'AFF_CHANTIER');
  if TOBL <> nil then
  begin
    if Trim(Copy(CodCha,2,14)) <> '' then   // correction BRL : certaines fois le code n'est pas à blanc mais à "A               00"
    begin
      SvgCodeAff := TobL.GetValue(prefixe+'_AFFAIRE');
      TobL.PutValue(prefixe+'_AFFAIRE',CodCha);
    end;
  end;
end;
{$ENDIF}
if TypeChargement = tCodeSection  then NomTT := 'GCCHAMPSANA' else
if TypeChargement = tLibSection   then NomTT := 'GCCHAMPSANA';
TobTab:=Tob.Create('_Tablettes',nil,-1);
if TOBGCS = nil then
begin
	QQ:=OpenSQL('Select CO_LIBELLE,CO_LIBRE from Commun where CO_TYPE="GCS" and CO_LIBRE Like "%'+VH_GC.GCMarcheVentilAna+'%"',True) ;
	if Not QQ.Eof then TobTab.LoadDetailDB('COMMUN','','',QQ,False,True) ;
end else
begin
	TOBTab.dupliquer (TOBGCS,true,true);
end;
Ferme (QQ) ;
for i:=0 to TobTab.Detail.count-1 do
  begin
  Champs:=TobTab.Detail[i].GetValue('CO_LIBRE') ;
  TobTab.Detail[i].AddChampSup('_CHAMPS',False) ;
  TobTab.Detail[i].PutValue('_CHAMPS',ReadTokenSt(Champs)) ;
  end ;
Champs:='' ;
for i:=0 to TOBStruc.Detail.Count-1 do
    BEGIN
    TOBDet:=TOBStruc.Detail[i];
    if (TypeChargement=tCodeSection) and (TOBDet.GetValue('GDA_TYPESTRUCTANA')='LSE') then continue ;
    if (TypeChargement=tLibSection) and (TOBDet.GetValue('GDA_TYPESTRUCTANA')='SEC') then continue ;
    Champs:=TOBDet.GetValue('GDA_LIBCHAMP') ;
    NomChamps:=Champs ;
    IsLibelle := TobDet.GetValue('GDA_ISLIBELLE') ;
    if Champs <> '' then
      BEGIN
      if Copy(Champs,1,9)='Constante' then
      begin
        Champs := TOBDet.getString('GDA_CSTTE');
      end else
      begin
        Champs := RecupChampsTob (Champs, TOBPiece,TOBL,TOBA,TOBTiers,nil);
        If IsLibelle='X' then Champs := RecupLibelle (NomChamps,Champs, TOBPiece,TOBL,TOBA,TOBTiers,TOBTab);
      end;
      Lng := StrToInt (TobDet.GetValue('GDA_LONGUEUR'));
      // on renseigne le CodeVide avec le caractère de bourrage pour comparer
      // avec le code section en fin de fonction
      CodeVide := CodeVide + BourreChamps ('',Lng,TFichierBase(Ord(fbAxe1)+NumA-1),True);
      // Bourrage effectué avec le caractère de bourrage uniquement pour le code
      // pour le libellé on met des espaces
      if Lng <> 0 then
        if TypeChargement = tLibSection then Champs:=BourreChamps (Champs,Lng,TFichierBase(Ord(fbAxe1)+NumA-1),False)
                                        else Champs:=BourreChamps (Champs,Lng,TFichierBase(Ord(fbAxe1)+NumA-1),True);
      CodeAna := CodeAna + Copy (Champs,1,Lng);
      END;
    END;
if TypeChargement = tCodeSection  then
  Begin
  // Si le code ne contient que des caractères de bourrage, on le vide pour
  // pouvoir utiliser la section par défaut
  if CodeAna = CodeVide then CodeAna := '';

  // Si le code est renseigné, on complète avec le caractère de bourrage
  if length(CodeAna)>0 then Result := BourreLaDonc(CodeAna,TFichierBase(Ord(fbAxe1)+NumA-1))
  else Result := '';
  End
else
begin
  // cadrage à gauche du libellé
  Result := Trim(CodeAna);
end;

// Traitement particuliers pour les appels d'interventions : BRL 041207
//    Restauration du code affaire dans la tob LIGNE
If (tobl <> nil) and (CodCha <> '') then
begin
    TobL.PutValue(prefixe+'_AFFAIRE',SvgCodeAff);
end;

TobTab.Free;
END;

Function RecupChampsTob (Champs : String; TOBPiece,TOBL,TOBA,TOBTiers,TOBConso : TOB) : String;
Var Prefixe,PrefixeL,StChamp : string;
    Posit : integer;
BEGIN
Result := '';
Posit := Pos('_',Champs);
Prefixe  := copy (Champs,1,Posit-1);
StChamp := copy (Champs,Posit+1,255);
if TOBL <> nil then prefixeL := GetPrefixeTable (TOBl);
if Prefixe ='BCO' then // Champs consommations
    Result := TobConso.GetValue(Champs) else
if Prefixe ='T' then // Champs Tiers
    Result := TobTiers.GetValue(Champs) else
if Prefixe = 'YTC' then // Champs complément tiers ressource
    Result := GetChampsComplementTiers(TobTiers.GetValue('T_TIERS'),Champs) else
{$IFDEF AFFAIRE}
if Prefixe = 'AFF' then // Champs Affaire
	 begin
    if TOBL <> nil then Result := GetChampsAffaire (TobL.GetValue(PrefixeL+'_AFFAIRE'),Champs)
    							 else Result := GetChampsAffaire (TobPiece.GetValue('GP_AFFAIRE'),Champs)
   end else
{$endif}
if Prefixe = 'GA' then // Champs Article
	 begin
		if TOBA <> nil then
    begin
    	Result := TobA.GetValue(Champs);
		end;
	 end else
if Prefixe = 'GL' then
    begin // Champs ligne
    if TOBL = nil then Exit;
    if PrefixeL = 'BOP' then
    begin
      Result := TobL.GetValue(PrefixeL+'_'+StChamp);
    end else
    begin
      Result := TobL.GetValue(Champs)
    end;
    end else
if Prefixe = 'BOP' then // Champs ligne ouvrage a plat
	  begin
    if TOBL = nil then Exit;
    Result := TobL.GetValue(Champs)
    end else
if Prefixe = 'GP' then // Champs pièce
    Result := Tobpiece.GetValue(Champs);
END;

Function RecupChampsTobGenerique (Champs : String; TOBPiece,TOBL,TOBAFF,TOBA,TOBTiers,TOBConso : TOB) : String;
Var Prefixe,PrefixeL,StChamp : string;
    Posit : integer;
BEGIN
Result := '';
Posit := Pos('_',Champs);
Prefixe  := copy (Champs,1,Posit-1);
StChamp := copy (Champs,Posit+1,255);
if TOBL <> nil then prefixeL := GetPrefixeTable (TOBl);
if Prefixe ='BCO' then // Champs consommations
    Result := TobConso.GetValue(Champs) else
if Prefixe ='T' then // Champs Tiers
    Result := TobTiers.GetValue(Champs) else
if Prefixe = 'YTC' then // Champs complément tiers ressource
    Result := GetChampsComplementTiers(TobTiers.GetValue('T_TIERS'),Champs) else
{$IFDEF AFFAIRE}
if Prefixe = 'AFF' then // Champs Affaire
	 begin
   	Result := GetChampsAffaire (TobAFF.GetValue('AFF_AFFAIRE'),Champs)
   end else
{$endif}
if Prefixe = 'GA' then // Champs Article
		if TOBA <> nil then
    begin
    	Result := TobA.GetValue(Champs)
		end else
if Prefixe = 'GL' then
    begin // Champs ligne
    if TOBL = nil then Exit;
    if PrefixeL = 'BOP' then
    begin
      Result := TobL.GetValue(PrefixeL+'_'+StChamp);
    end else
    begin
      Result := TobL.GetValue(Champs)
    end;
    end else
if Prefixe = 'BOP' then // Champs ligne ouvrage a plat
	  begin
    if TOBL = nil then Exit;
    Result := TobL.GetValue(Champs)
    end else
if Prefixe = 'GP' then // Champs pièce
    Result := Tobpiece.GetValue(Champs);
END;

Function RecupLibelle (NomChamps,Champs : String; TOBPiece,TOBL,TOBA,TOBTiers,TobTab : TOB) : String;
Var Tab : string;
    TobT: Tob ;
BEGIN
  Result := '';
if Champs = '' then Exit;
TobT:=TobTab.FindFirst(['_CHAMPS'],[NomChamps],False) ;
if  TobT<>nil then
  begin
  // MODIF BTP 14/03/03 :
  // Pour lire l'enregistrement correspondant au libellé s'il est placé
  // après le code dans la tablette GCCHAMPSANA
  // ex : pour famille comptable article, code avec CO_CODE=GFC et libellé avec CO_CODE=GFL
  Tab:=TobT.GetValue('CO_LIBRE') ;
  ReadTokenSt(Tab) ;
  if Copy(Tab,1,1)<>'X' then TobT:=TobTab.FindNext(['_CHAMPS'],[NomChamps],False) ;
  //---
  Tab:=TobT.GetValue('CO_LIBELLE') ;
  ReadTokenSt(Tab) ;
  Tab:=ReadTokenSt(Tab);
  end ;
if Tab<>'' then Result:=RechDom(Tab,Champs,False)
else Result:=Champs ;
END;


Function CreatLesSousSections (NumAxe : Integer; Cpt: string ; TobStrucCode,TOBPiece,TOBL,TOBA,TOBTiers : TOB) : Boolean;
Var i, Posit, Lng  : integer;
    TobDet : TOB;
    CodeTT, PlanSS,Code,Libelle : String;
    TraiteSS : Boolean;
BEGIN
Result := True ;
Code := ''; Libelle := '';
for i:=0 to TOBStrucCode.Detail.Count-1 do
    BEGIN
    TOBDet:=TOBStrucCode.Detail[i];
    if (TobDet.GetValue('GDA_CREATSSSECTION')= 'X') then TraiteSS := True else TraiteSS := False;
    if Not(TraiteSS) then Continue;
    CodeTT := TobDet.GetValue('GDA_LIBCHAMP');
    Posit := strToInt(TobDet.GetValue('GDA_POSITION'));
    Lng := StrToInt (TobDet.GetValue('GDA_LONGUEUR'));
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

Function BourreChamps (Champs: String; Lng:Integer;LeType : TFichierBase; CarBour:boolean): String ;
Var i,Lch: Integer ;
Bourre: String ;
begin
Lch := Length(Champs) ;
If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   BEGIN
   If LeType=fbNatCpt Then Bourre:=VH^.BourreLibre Else Bourre:=VH^.Cpta[LeType].Cb ;
   if not(CarBour) then Bourre:=' '; //mcd 15/01/2003 
   If Bourre<>'' then  //mcd test '' au lieu ' '
      If Lch<Lng then for i:=Lch+1 to Lng do Champs:=Champs+Bourre ;
   END ;
Result:=Champs ;
end ;

end.


