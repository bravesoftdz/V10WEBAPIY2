unit UtilDispGC;

interface

Uses windows,hent1,entGC,hctrls, sysutils,math, AglInit,HmsgBox,
{$IFDEF EAGLCLIENT}
     MenuOLX,Maineagl,
{$ELSE}
     MenuOLG,Fe_Main,
{$ENDIF}
{$IFDEF EDI}
     EntEDI,
{$ENDIF}
		 LicUtil,
     UApplication,
     ParamSoc,UtilArticle,utilpgi,Ent1;

{$IFDEF NOMADE}
  procedure ChargeModules_PCP;
{$ENDIF}
	Procedure AfterProtecGC ( sAcces : String ) ;
  procedure ChargeModules ( Charge : Boolean ) ;
  procedure AjusteMenu ( NumModule : integer ) ;
  Procedure DispatchTTArticle ( Action : TActionFiche ; Lequel,TT,Range : String ) ;
  {Spécificités Modules, CEGID, ...}
  Procedure TraiteSpecifsCEGID ( NumModule : integer ) ;
  Procedure TraiteMenusAffaire ( NumModule : integer ) ;
  Procedure TraiteMenusGRC ( NumGroupe : integer; bNumModule : boolean = true ) ;
  Procedure TraiteMenus303265 ( NumModule : integer ) ;   // pour compatibilité avec l'antérieur
  Procedure TraiteMenusGescom ( NumModule : integer ) ;
  Procedure TraiteMenusEAGL ( NumModule : integer ) ;
  Procedure TraiteMenusMode ( NumModule : integer ) ;
  Procedure TraiteMenusModeEAGL ( NumModule : integer ) ;
  Function  TraiteMenuSpecif ( NumModule : integer ) : boolean;
  Procedure TripoteMenuGCS3 ( NumModule : integer ) ;
{$IFDEF NOMADE}
  Procedure TraiteMenusNomade ( NumModule : integer);
  procedure SupprimeDansMenu (iIndDeb, iIndFin : integer; bItem : boolean);
{$ENDIF NOMADE}
  Procedure TraiteMenusTarifs;
  Procedure TraiteMenusGRCModule ( NumModule : integer ) ;
Const
(*
{$IFDEF BUSINESSPLACE}
  GCCodeDomaine = '00086010' ;
  GCCodesSeria : array[1..9] of string =('00083090','00311090','00312090','00313090','00314090','00316090','00416090','14900090','14902090');
  GCTitresSeria : array[1..9] of Hstring =('BUSINESS PLACE BTP : Etudes & Situations','BUSINESS PLACE BTP : Récupération des Appels d''offres','BUSINESS PLACE BTP : Achats/Stocks','BUSINESS PLACE BTP : Gestion des chantiers','BUSINESS PLACE BTP : Relations Clients','BUSINESS PLACE BTP : Contrats','BUSINESS PLACE BTP : Interventions','BUSINESS PLACE BTP : Cotraitance','BUSINESS PLACE BTP : Sous-traitance');
{$ELSE}
    //uniquement en line
{*  Ent_Message = 'Business Line BTP : ';
    GCCodeDomaine = '00595010' ;
    GCCodesSeria : array[1..2] of string =('00594090', '00648090');
    GCTitresSeria : array[1..2] of Hstring =(Ent_Message + 'Devis-Factures', Ent_Message + 'Gestion Chantiers');
*}
    GCCodeDomaine = '00056010' ;
    GCCodesSeria : array[1..9] of string =('00055090','00084090','00094090','00210090','00254090','00255090','00256090','14901090','14903090');
    GCTitresSeria : array[1..9] of Hstring =('BUSINESS SUITE BTP : Etudes & Situations','BUSINESS SUITE BTP : Récupération des Appels d''offres','BUSINESS SUITE BTP : Achats/Stocks','BUSINESS SUITE BTP : Gestion des chantiers','BUSINESS SUITE BTP : Relations Clients','BUSINESS SUITE BTP : Contrats','BUSINESS SUITE BTP : Interventions','BUSINESS SUITE : Cotraitance','BUSINESS SUITE : Sous-traitance');
{$ENDIF}
*)
{$IFDEF NOMADE}
    GCModules:array[0..4] of integer =(30,31,32,169,171);
    GCImages :array[0..4] of integer =(35,18,5,69,78);
{$ELSE NOMADE}
    {$IFDEF GPAO}
      GCModules: Array[0..{$IFNDEF EDI}14{$ELSE}15{$ENDIF}] of integer = (210, 211, 212, 213, 70, 59, 92, 215, 120, 126, 122, 121{$IFDEF EDI}, 260{$ENDIF EDI}, 125,  98, 219);
      GCImages : Array[0..{$IFNDEF EDI}14{$ELSE}15{$ENDIF}] of Integer = ( 35,  18,   5,  11,  8, 32, 77,  28,  60,  51, 121,  26{$IFDEF EDI},  67{$ENDIF EDI},  36, 106,  49);
    {$ELSE GPAO}
      {$IFDEF GESCOM}
        {$IFDEF GPAOLIGHT}
        GCModules:array[0..13] of integer =(30,36,31,32,33,59,70,92,160,111,267,260,65,60);
        GCImages :array[0..13] of integer =(35,41,18, 5,11,32, 8,77, 60, 69, 59, 67,28,49);
        {$ELSE GPAOLIGHT}
        GCModules:array[0..12] of integer =(30,36,31,32,33,59,70,92,160,111,260,65,60);
        GCImages :array[0..12] of integer =(35,41,18, 5,11,32, 8,77, 60, 69, 67,28,49);
        {$ENDIF GPAOLIGHT}
      {$ELSE GESCOM}
        GCModules:array[0..9] of integer =(30,31,32,33,59,70,92,65,60,160);
        GCImages :array[0..9] of integer  =(35,18, 5,11,32, 8,77,28,49,60);
      {$ENDIF GESCOM}
    {$ENDIF GPAO}
{$ENDIF NOMADE}

ModulesCommuns:array[0..6] of integer =(30,31,32,33,65,60,160) ;
ImagesCommuns :array[0..6] of integer =(35,18, 5,11,28,49,60) ;

ModuleGRC        :array[0..0] of integer =(92);  ImageGRC        :array[0..0] of integer =(77);
ModuleECommerce  :array[0..0] of integer =(59);  ImageECommerce  :array[0..0] of integer =(32);
ModuleAffaire    :array[0..0] of integer =(70);  ImageAffaire    :array[0..0] of integer =(8);
ModuleAssemblage :array[0..0] of integer =(267); ImageAssemblage :array[0..0] of integer =(59);
ModulePCPServeur :array[0..0] of integer =(111); ImagePCPServeur :array[0..0] of integer =(69);
ModuleEDI        :array[0..0] of integer =(260); ImageEDI        :array[0..0] of integer =(67);

implementation
uses UFonctionsCBP;

Function estDans( ii : integer; const AA : array of integer) : boolean;
var i : integer;
begin
Result:=false;
For i:=0 to High(AA) do Result:=((AA[i]=ii) or Result);
end;

Procedure ChangeTitreOuVireMenu ( NumMenu : integer; Tablette, Code : string ) ;
Var ch : string;
BEGIN
ch:=RechDom(Tablette,Code,FALSE) ;
if Length(ch)> 0 then
   BEGIN
   if (copy(ch,1,2)='.-') then FMenuG.RenameItem(NumMenu,'-')
                          else FMenuG.RenameItem(NumMenu,ch);
   END;
END;

procedure ChargeModules ( Charge : Boolean ) ;
var  LesModules,LesImages : array of integer;
     i,j,HighModules:integer;
     OkOk:Boolean;
begin
HighModules:=High(GCModules);
{$IFDEF NOMADE}
if ctxMode in V_PGI.PGIContexte then
    BEGIN
    // DCA - Ajout des stocks dans PCP
    HighModules:=4;
    GCModules[0]:=167; GCModules[1]:=170; GCModules[2]:=168; GCModules[3]:=169; GCModules[4]:=171;
    GCImages[0] :=94;  GCImages[1] :=121; GCImages[2] :=1;   GCImages[3] :=69;  GCImages[4] :=78;
    END ;
{$ENDIF}
SetLength(LesModules,HighModules+1);
SetLength(LesImages,HighModules+1);
J:=0;
for i:= 0 to HighModules do
    begin
{$IFDEF NOMADE}
    OkOK := True ;
{$ELSE}
    OkOK := False ;
    If estDans(GCModules[i],ModulesCommuns) then OkOk:=True
    else If estDans(GCModules[i],ModuleGRC) and (VH_GC.GRCSeria) then OkOk:=True
{$IFDEF CCS3}
    {$IFDEF PREMIUM}
    else If estDans(GCModules[i],ModuleAchatStock) and (VH_GC.GCAchatStockSeria)  then OkOk:=True
    else If estDans(GCModules[i],ModuleArticles)   and (not VH_GC.GCAchatStockSeria)  then OkOk:=True
    {$ENDIF PREMIUM}
{$ELSE CCS3}
    else If estDans(GCModules[i],ModuleECommerce)  and (VH_GC.ECSeria)         then OkOk:=True
{$IFDEF GPAOLIGHT}
    else If estDans(GCModules[i],ModuleAssemblage) and (VH_GC.OASeria)         then OkOk:=True
{$ENDIF}
{$IFDEF EDI}
    else If estDans(GCModules[i],ModuleEDI)        and (VH_EDI.EDISeria)        then OkOk:=True
{$ENDIF}
    else If estDans(GCModules[i],ModulePCPServeur) and (VH_GC.PCPServeurSeria) then OkOk:=True
    else If estDans(GCModules[i],ModuleAffaire)                                then OkOk:=True
{$ENDIF CCS3}

{$IFDEF GESCOM}
    else If estDans(GCModules[i],ModuleVteComptoir) and (VH_GC.VTCSeria) then OkOk:=True
{$ENDIF GESCOM}
    ;
    {$IFDEF EAGLCLIENT}
    if (GCModules[i]=60) and (not V_PGI.Superviseur) then Okok:=False ; //administration
    {$ENDIF}
{$ENDIF}
    if OkOk then
       BEGIN
       LesModules[j]:=GCModules[i] ; LesImages[j]:=GCImages[i] ;
       inc(J) ;
       END ;
    end;
SetLength(LesModules,J);
SetLength(LesImages,J);
FMenuG.SetModules(LesModules,LesImages) ;
if J > 5 then begin V_PGI.NbColModuleButtons:=2; V_PGI.NbRowModuleButtons:= ceil(J/2) ;end
         else begin V_PGI.NbColModuleButtons:=1 ; end;
end;

Procedure AfterProtecGC ( sAcces : String ) ;
var
{$IFDEF NOMADE}
    PCPSeria : boolean ;
{$ELSE}
    i : integer ;
{$ENDIF}

       Procedure GCSerialise(CodeSeria : string; bSeria : boolean) ;
          begin
          if (CodeSeria='00100') or (CodeSeria='00012') {S3} then  VH_GC.GCSeria :=bSeria else
          if (CodeSeria='00013') then  VH_GC.ECSeria :=bSeria else
          if (CodeSeria='05003') or (CodeSeria='00060') then  VH_GC.GRCSeria:=bSeria else
          if (CodeSeria='00099') then  VH_GC.GASeria :=bSeria else
          if (CodeSeria='00050') or (CodeSeria='00051') then  VH_GC.VTCSeria :=bSeria else
          if (CodeSeria='00032') then  VH_GC.GAPlanningSeria :=bSeria else
          if (CodeSeria='00129') then  VH_GC.OASeria :=bSeria else
          if (CodeSeria='00079') then  VH_GC.PCPServeurSeria :=bSeria else
{$IFDEF EDI}
          if (CodeSeria='00146') then  VH_EDI.EDISeria :=bSeria else
          if (CodeSeria='00139') then  VH_EDI.EDISeria_FAC :=(bSeria AND VH_EDI.EDISeria) else
          if (CodeSeria='00138') then  VH_EDI.EDISeria_CDE :=(bSeria AND VH_EDI.EDISeria) else
          if (CodeSeria='00137') then  VH_EDI.EDISeria_ARC :=(bSeria AND VH_EDI.EDISeria) else
          if (CodeSeria='00130') then  VH_EDI.EDISeria_BL  :=(bSeria AND VH_EDI.EDISeria) else
{$ENDIF}
{$IFDEF PREMIUM}
          if (CodeSeria='00106') then  VH_GC.GCAchatStockSeria :=bSeria ;
{$ENDIF}
          end;
BEGIN
VH_GC.GCSeria :=False; VH_GC.ECSeria  :=False ;
VH_GC.GRCSeria:=False; VH_GC.BOSeria  :=False;
VH_GC.GASeria :=False; VH_GC.VTCSeria :=False;
VH_GC.OASeria :=False; VH_GC.PCPServeurSeria := False;
VH_GC.GAPlanningSeria   := false;
{$IFDEF EDI}
VH_EDI.EDISeria := False;
VH_EDI.EDISeria_FAC := False; VH_EDI.EDISeria_CDE := False;
VH_EDI.EDISeria_ARC := False; VH_EDI.EDISeria_BL  := False;
{$ENDIF}
{$IFDEF PREMIUM}
{$IFDEF CCS3}
VH_GC.GCAchatStockSeria := false;
{$ELSE}
VH_GC.GCAchatStockSeria := True;
{$ENDIF}
{$ENDIF}

{$IFDEF NOMADE}
if Length(sAcces) > 0 then PCPSeria := (sAcces[1] = 'X')  // PCP sérialisé
else PCPSeria := False ;

if V_PGI.NoProtec then
   begin
   V_PGI.VersionDemo := True ;
   PCPSeria:=False ;
   end
   else V_PGI.VersionDemo := False ;

VH_GC.PCPVenteSeria := PCPSeria;
if (PCPSeria=False) then  V_PGI.VersionDemo:= True; // on repasse en demo dans ce cas
if (V_PGI.VersionDemo) and (Pos(' (DEMO)',TitreHalley)=0) then TitreHalley:=TitreHalley+' (DEMO)' ;

{$ELSE}

//Verification des options sérialisées
if Length(sAcces)>0 then for i:=1 to Length(sAcces) do GCSerialise(copy(GCCodesSeria[i],1,5), (sAcces[i]='X') ) ;
  {$IFNDEF GPAO}
    V_PGI.VersionDemo:=((not VH_GC.GCSeria) or V_PGI.NoProtec) ;
  {$ENDIF GPAO}
{if VH_GC.GCIfDefCEGID then
begin
  V_PGI.VersionDemo:=False;
  VH_GC.GCSeria:=True  ; VH_GC.GRCSeria:=True ; VH_GC.GASeria :=True ;
  VH_GC.ECSeria :=True  ;
end;
}
{$ENDIF}

{$IFDEF GRC}
if (V_PGI.VersionDemo) then VH_GC.GRCSeria:=True;
if VH_GC.GRCSeria then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC]
else if (ctxGRC in V_PGI.PGIContexte) then V_PGI.PGIContexte:=V_PGI.PGIContexte-[ctxGRC]  ;
if (ctxGRC in V_PGI.PGIContexte) then
   V_PGI.TabletteHierarchiques:=true;
{$ENDIF}

{$IFDEF GPAO}
if ctxGpao in V_PGI.PGIContexte then
begin
    {$IFDEF GRC}
    V_PGI.PGIContexte := V_PGI.PGIContexte+[ctxGRC];
    {$ENDIF GRC}
    V_PGI.CodeProduit:='034';     // V500_001 KB 28/10/2003 Initialisation variable pour WebSat
    AfterProtecGpao (sAcces);
    exit;
end;
{$ENDIF GPAO}

{$IFDEF AFFAIRE}
{$IFNDEF CCS3}
if (V_PGI.VersionDemo) then VH_GC.GASeria:=True;
VH_GC.GAPlanningSeria:=False;  //gm 13/01/03
{$ENDIF}
{$ENDIF}
{$IFDEF GESCOM}
if (V_PGI.VersionDemo) then
begin
  VH_GC.VTCSeria:=True;
  {$IFDEF GPAOLIGHT} VH_GC.OASeria:=True; {$ENDIF}
  {$IFDEF NOMADESERVER} VH_GC.PCPServeurSeria:=True; {$ENDIF}
  {$IFDEF EDI}
  VH_EDI.EDISeria:=true;
  VH_EDI.EDISeria_FAC := true; VH_EDI.EDISeria_CDE := true;
  VH_EDI.EDISeria_ARC := true; VH_EDI.EDISeria_BL  := true;
  {$ENDIF}
end;
{$ENDIF}

{$IFDEF PREMIUM}
if (V_PGI.VersionDemo) then VH_GC.GCAchatStockSeria:=True;
{$ENDIF}


VH_GC.GAAchatSeria:=True; // toujours vrai en GC

ChargeModules(True) ;
// Initialisation variable pour WebSat
V_PGI.CodeProduit:='010'; // GC
if VH_GC.GRCSeria then V_PGI.CodeProduit:=V_PGI.CodeProduit+';012' ; // GRC
END ;

{$IFDEF NOMADE}
{************************************************************
Auteur  ...... : LS
Créé le ...... : 04/12/2003
Modifié le ... :
Description .. : Gestion des menus PCP
Description .. : La procédure ChargeModules n'est pas du tout utilisée pour PCP
Description .. :                    GC   MODE
                 Vente............. 279  167
                 Achat............. 280  171
                 Article........... 281  -
                 Stock............. -    170
                 Données de base... -    168
                 Transmissions..... 169  169
*****************************************************************}
procedure ChargeModules_PCP;
begin
//  GetAchatVente;
  GCModules[0] := -1; GCImages[0] := -1;
  GCModules[1] := -1; GCImages[1] := -1;
  GCModules[2] := -1; GCImages[2] := -1;
  GCModules[3] := -1; GCImages[3] := -1;
  GCModules[4] := -1; GCImages[4] := -1;
  if (VH_GC.PCPUsVte) and (VH_GC.PCPUsAch) then // Vente et Achat
  begin
    GCModules[0] := 285; GCImages[0] := 35;
//    GCModules[1] := 280; GCImages[1] := 18;
    GCModules[1] := 149; GCImages[1] := 5;
    GCModules[2] := 169; GCImages[2] := 69;
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    begin
    	GCModules[3] := 148; GCImages[3] := 104;
    	GCModules[4] := 60; GCImages[4] := 34;
    end;
  end else
  if VH_GC.PCPUsVte then // Vente
  begin
    GCModules[0] := 285; GCImages[0] := 35;
    GCModules[1] := 149; GCImages[1] := 5;
    GCModules[2] := 169; GCImages[2] := 69;
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    begin
    	GCModules[3] := 148; GCImages[3] := 104;
    	GCModules[4] := 60; GCImages[4] := 34;
    end;
  end else
  if VH_GC.PCPUsAch then // Achat
  begin
//    GCModules[0] := 280; GCImages[0] := 18;
//    GCModules[1] := 281; GCImages[1] := 5;
//    GCModules[2] := 169; GCImages[2] := 69;
  end else
  begin // Aucun des 2
    GCModules[0] := 149; GCImages[0] := 5;
    GCModules[1] := 169; GCImages[1] := 69;
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    begin
    	GCModules[2] := 60; GCImages[2] := 34;
    end;
  end;
  FMenuG.SetModules(GCModules,GCImages) ;
end;
{$ENDIF NOMADE}

Procedure TraiteMenusMode ( NumModule : integer ) ;
BEGIN
END ;

Procedure TraiteMenusModeEAGL ( NumModule : integer ) ;
BEGIN
END ;


Procedure TraiteMenusEAGL ( NumModule : integer ) ;
BEGIN
{$IFDEF EAGLCLIENT}
Case NumModule of
    30 : Begin
         FMenuG.RemoveItem(30212);// epuration des pieces
         FMenuG.RemoveItem(30507);// Suppression clients
         //Affaire
         FMenuG.RemoveItem(30414);
         //edition // DBR : Donner accès a ces deux états - Fiche qualité
//         FMenuG.RemoveItem(30602);   //traites
//         FMenuG.RemoveItem(30604);   //commissions représentants
         //Tarifs
//         FMenuG.RemoveItem(30903);   //tarif TTC categorie client /artcle // DBR Fiche 10096
//         FMenuG.RemoveItem(30904);   //tarif TTC client /artcle // DBR Fiche 10096
         end;
    31 : BEGIN
         FMenuG.RemoveGroup(31300,True) ;
         FMenuG.RemoveItem(31205); // epuration des pieces
         FMenuG.RemoveItem(31705); // fermeture fournisseurs
         FMenuG.RemoveItem(31706); // ouverture fournisseurs
         FMenuG.RemoveItem(31707); // supression fournisseurs
         END ;
    32 : BEGIN
         FMenuG.RemoveGroup(32700,True) ; // dimensions
         FMenuG.RemoveItem(32203);// fin de mois
         FMenuG.RemoveItem(32204);// fin d'exercice
         FMenuG.RemoveItem(32206);// initialisation des stocks
//         if not VH_GC.GCIfDefCEGID then // DBR Fiche 10756
//            FMenuG.RemoveGroup(32300,True) ; // inventaires // DBR Fiche 10756
         //FMenuG.RemoveItem(32201);// mvts interdepots
         FMenuG.RemoveItem(32405);// Consommations
         END ;
    33 : BEGIN
         // affaire
         FMenuG.RemoveGroup(33400,True); FMenuG.RemoveItem(33130);
         FMenuG.RemoveItem(33131);
         if not VH_GC.GCIfDefCEGID then
           FMenuG.RemoveGroup(33140,True); // stats Eric
         //FMenuG.RemoveItem(33132);     // VentesCube Affaire
         FMenuG.RemoveItem(33201);     // Analyse Derniers prix d'achats
         FMenuG.RemoveItem(33111); // Graphe    (plantage)
         FMenuG.RemoveItem(33121); // Graphe   (plantage)
         FMenuG.RemoveItem(33211); // Graphe   (plantage)
         FMenuG.RemoveItem(33221); // Graphe   (plantage)
         FMenuG.RemoveItem(33501);     // DEB
         END ;
    36 : BEGIN
         // ventes comptoir
         FMenuG.RemoveGroup(36800,True); // Modèles d'impression
         FMenuG.RemoveItem(36901); // Paramètrage caisse
         FMenuG.RemoveItem(36906); // Détail des éspèces
         FMenuG.RemoveItem(36902); // Modification en série
         END;
    59 : BEGIN
         // e-commerce
         FMenuG.RemoveItem(59101);  FMenuG.RemoveItem(59102);
         FMenuG.RemoveItem(59103); FMenuG.RemoveItem(59201); FMenuG.RemoveItem(59203);
         FMenuG.RemoveItem(59301); FMenuG.RemoveItem(59302); FMenuG.RemoveItem(59304);
         FMenuG.RemoveItem(59305); FMenuG.RemoveItem(59401);
         // affaire
         FMenuG.RemoveItem(59204); FMenuG.RemoveItem(59205);
         FMenuG.RemoveItem(59206); FMenuG.RemoveItem(59207);
         FMenuG.RemoveItem(59208); FMenuG.RemoveItem(59209);
         END;
    60 : BEGIN
         // administration
         FMenuG.RemoveGroup(60100,True); // Societe
         FMenuG.RemoveGroup(60400,True); // Traitement
         FMenuG.RemoveItem(3172);  FMenuG.RemoveItem(33221);
         FMenuG.RemoveItem(60203); FMenuG.RemoveItem(60204); FMenuG.RemoveItem(60205);
         END;
    65 : BEGIN
         // Paramètres
            // AGL
         FMenuG.RemoveItem(65905);   //region
         FMenuG.RemoveItem(65906);   //Code postaux
         FMenuG.RemoveGroup(65200,True);  // Société
            // param comptable
         FMenuG.RemoveItem(65907);  //Devise
         FMenuG.RemoveItem(65908);  //Regimes fiscaux
         FMenuG.RemoveItem(-65910); // tva et TPF
            //autres
         FMenuG.RemoveItem(65402);   //souches
         FMenuG.RemoveItem(65405);   //Arrondi
            // Affaires
         FMenuG.RemoveItem(65422);   //Blocage affaire
         FMenuG.RemoveItem(65423);   //Confidentialité
         FMenuG.RemoveItem(65424);   //Confidentialité
         END;
    70 : BEGIN
         // affaire
         FMenuG.RemoveItem(70562);
         //FMenuG.RemoveGroup(70580,True); // traitements ressource
         END;
   END ;
{$ENDIF}
END ;

Procedure TraiteMenus303265 ( NumModule : integer ) ;  // pour compatibilité avec l'antérieur
begin
 TraiteMenusGescom (NumModule) ;
end;

{$IFDEF NOMADE} //Enlève les icônes non utilisées
procedure SupprimeDansMenu (iIndDeb, iIndFin : integer; bItem : boolean);
var iInd : integer;
begin
for iInd := iIndDeb to iIndFin do
    begin
    if bItem then FMenuG.RemoveItem (iInd)
    else FMenuG.RemoveGroup (iInd, true);
    end;
end;

Procedure TraiteMenusNomade(NumModule : integer);
BEGIN
END ;
{$ENDIF} //NOMADE

Procedure TraiteMenusGescom ( NumModule : integer ) ;
Var i : integer ;
BEGIN
Case NumModule of
   30 : BEGIN
        {$IFNDEF PASREL}
        FMenuG.RemoveItem(30406) ;  {Pas.Rel}
        {$ENDIF}
         // Libellés des avoirs
        FMenuG.RenameItem(30261,GetInfoParPiece('AVC','GPP_LIBELLE'));
        FMenuG.RenameItem(30262,GetInfoParPiece('AVS','GPP_LIBELLE'));
        FMenuG.RenameItem(30236,GetInfoParPiece('AVC','GPP_LIBELLE'));
        FMenuG.RenameItem(30238,GetInfoParPiece('AVS','GPP_LIBELLE'));

        END ;
  32 :  begin
        { Categories de dimensions }
        for i:=1 to 5 do FMenuG.RenameItem(32710+i,RechDom('GCCATEGORIEDIM','DI'+IntToStr(i),FALSE)) ;
        {$IFDEF CCS3}
        FMenuG.RemoveGroup(32700,True); // Dimensions
        FMenuG.RemoveItem(32602);  // Emplacements
        {$ENDIF}
          if not GetParamSoc('SO_GCTRV') then
            FMenuG.RemoveGroup(21220, True)   //Logistique
          else
          begin
            if V_PGI.NumVersionSoc > 614 then
            begin
              if not GetParamSoc('SO_TRANSFPROP') then
                FMenuG.RemoveItem(-21221);       //Validation des propositions
              if not GetParamSoc('SO_TRANSFSCAN') then
              begin
                FMenuG.RemoveItem(-21222);       //Scan
                {$IFDEF GPAO}
                  FMenuG.RemoveItem(-21224);         //Edition inter-dépôts
                {$ELSE GPAO}
                  FMenuG.RemoveItem(-21225);         //Edition inter-dépôts
                {$ENDIF GPAO}
              end;
            end;
          end;
        end ;
  33 :  begin
        {$IFDEF CCS3}
        FMenuG.RemoveGroup(33300,True); // Traçabilité
        FMenuG.RemoveItem(-33130);       // Analyses par affaires
        {$ENDIF}
        end ;
  36 :  begin // DBR - Droits d'acces de la facturation comptoir qui n'est pas un menu
        FMenuG.RemoveGroup (36600, True);
        end;
  60 :  BEGIN
        {Champs obligatoire et restrictions fiches}
        {$IFDEF CCS3}
        FMenuG.RemoveItem(60206);  // Champs obligatoires
        FMenuG.RemoveItem(60207);  // restrictions fiches
        {$ENDIF}
        end ;
  65 :  BEGIN
        { Libellé des codes famille }
        for i:=1 to 3 do FMenuG.RenameItem(65320+i,RechDom('GCLIBFAMILLE','LF'+IntToStr(i),FALSE)) ;
        AvertirTable ('GCZONELIBREART');
        { Libellé des zones libres articles }
        for i:=1 to 9 do FMenuG.RenameItem(65500+i,RechDom('GCZONELIBREART','AT'+IntToStr(i),FALSE)) ;
        FMenuG.RenameItem(65510,RechDom('GCZONELIBREART','ATA',FALSE)) ;
        { Libellé des zones libres tiers }
        for i:=1 to 9 do FMenuG.RenameItem(65510+i,RechDom('GCZONELIBRETIE','CT'+IntToStr(i),FALSE)) ;
        FMenuG.RenameItem(65520,RechDom('GCZONELIBRETIE','CTA',FALSE)) ;
        { Libellé des zones libres établissements }
        for i:=1 to 9 do FMenuG.RenameItem(65550+i,RechDom('GCZONELIBRE','ET'+IntToStr(i),FALSE)) ;
        FMenuG.RenameItem(65560,RechDom('GCZONELIBRE','ETA',FALSE)) ;
        for i:=1 to 3 do FMenuG.RenameItem(65544+i,RechDom('GCZONELIBRETIE','FT'+IntToStr(i),FALSE)) ;
        { Libellé des zones libres Contacts }
//        for i:=1 to 3 do FMenuG.RenameItem(65565+i,RechDom('GCZONELIBRE','BT'+IntToStr(i),FALSE)) ;
        for i:=1 to 9 do FMenuG.RenameItem(65670+i,RechDom('GCZONELIBRE','BT'+IntToStr(i),FALSE)) ;
        FMenuG.RenameItem(65680,RechDom('GCZONELIBRE','BTA',FALSE)) ;

{$IFNDEF CCS3}
        FMenuG.RemoveItem(65657);  // départements pour S3
{$ENDIF}
{$IFDEF CCS3}
        FMenuG.RemoveItem(65204);  // assistant code affaire
        FMenuG.RemoveItem(65408);  // Ventil ana par affaire
        FMenuG.RemoveItem(65305);  // types emplacements
        FMenuG.RemoveItem(65306);  // cotes emplacements
        // affaires
        FMenuG.RemoveItem(-65524);  // Titre libre affaire
        FMenuG.RemoveItem(-65250);  // AFFAIRE
        FMenuG.RemoveGroup(65600,True); // Tablettes affaire
{$ENDIF}
        END  ;
 160 :  BEGIN
        {$IFDEF CCS3}
        FMenuG.RemoveItem(160101);  // Etats libres
        {$ENDIF}
       END;
     end ;
END ;

Procedure TraiteMenusTarifs;
var
  lPrefSystTarif : boolean;
begin
  lPrefSystTarif := GetParamSoc('SO_PREFSYSTTARIF');
  FMenuG.RemoveGroup(30100, lPrefSystTarif    ); //Suppression Tarifs clients H.T.
  FMenuG.RemoveGroup(30900, lPrefSystTarif    ); //Suppression Tarifs clients T.T.C.
  FMenuG.RemoveGroup(2104 , not lPrefSystTarif); //Suppression Tarifs clients
  FMenuG.RemoveGroup(31400, lPrefSystTarif    ); //Suppression Tarifs Fournisseurs
  FMenuG.RemoveGroup(2114 , not lPrefSystTarif); //Suppression Tarifs Fournisseurs
end;

Procedure TraiteMenusAffaire ( NumModule : integer ) ;
BEGIN
Case NumModule of
   30 : BEGIN
        if Not VH_GC.GASeria then
           begin
           FMenuG.RemoveItem(30210); FMenuG.RemoveItem(30225); FMenuG.RemoveItem(30240);
           FMenuG.RemoveItem(30414); FMenuG.RemoveItem(30407); FMenuG.RemoveItem(30408);
           FMenuG.RemoveItem(30263); FMenuG.RemoveItem(30237);
           FMenuG.RemoveGroup(30800,True); // editions affaires
           end;
        END ;
   33 : BEGIN
        if Not(VH_GC.GASeria) then
           begin
           FMenuG.RemoveGroup(33400,True);
           FMenuG.RemoveGroup(-33130,True);
           end;
        END ;
   65 : BEGIN
        if Not VH_GC.GASeria then
        begin
          FMenuG.RemoveItem(-65420);
          FMenuG.RemoveItem(65654); FMenuG.RemoveItem(65655); FMenuG.RemoveItem(65656);
        end else
        begin
          if not GetParamSoc('SO_AFAPPCUTOFF') then FMenuG.RemoveItem(65425);
          if not GetParamSoc('SO_AFLIENPAIEANA') and not GetPAramSoc('SO_AFLIENPAIEVAR') then
            FMenuG.RemoveItem(65426);
        end;
        END ;
   70 : BEGIN
        if Not VH_GC.GASeria then
        begin
          FMenuG.RemoveItem(70101); FMenuG.RemoveItem(-70110);
          FMenuG.RemoveItem(70106); FMenuG.RemoveItem(70230);
          FMenuG.RemoveItem(70440);
          FMenuG.RemoveGroup(70300,True); FMenuG.RemoveGroup(70550,True);
          FMenuG.RemoveGroup(70400,True); FMenuG.RemoveGroup(70500,True);
          FMenuG.RemoveGroup(70600,True); FMenuG.RemoveGroup(70700,True);
          FMenuG.RemoveGroup(-70560,True);
        end else
        begin
          FMenuG.RemoveItem(70280); // cacher le suivi des propositions missions
          if (GetParamSoc('SO_AFVISAACTIVITE')<>True) then FMenuG.RemoveItem(70552);
          if not(GetParamSoc('SO_PgLienRessource')) or (GetParamSoc('SO_AFLIENPAIEDEC')) then
            FMenuG.RemoveItem(70585); //Affaire - Lien Paie - ressource/salarié
          if not GetParamSoc('SO_AFAPPCUTOFF') then
            FMenuG.Removegroup(-70560,True);   // pas de  cutoff
          if not GetParamSoc('SO_AFLIENPAIEANA') and not GetPAramSoc('SO_AFLIENPAIEVAR') then
          begin
            FMenuG.RemoveItem(70570);
            FMenuG.RemoveItem(70571);
          end;
        end;
        END ;
   end ;
END ;

Procedure TraiteMenusGRC ( NumGroupe : integer; bNumModule : boolean = true ) ;
Var i : integer ;
    St : String ;

  procedure FMenuGRemoveItem (NumTag : integer);
  begin
    { GC_DBR_20081023_FQ;034;15656 }
    if (NumGroupe = -1) or bNumModule or FMenuG.OptionAppartientGroupe (NumGroupe, NumTag) then
    begin
      FMenuG.RemoveItem (NumTag);
    end;
  end;

  procedure FMenuGRenameItem (NumTag : integer; Libelle : string);
  begin
    { GC_DBR_20081023_FQ;034;15656 }
    if (NumGroupe = -1) or bNumModule or FMenuG.OptionAppartientGroupe (NumGroupe, NumTag) then
    begin
      FMenuG.RenameItem (NumTag, Libelle);
    end;
  end;

  function FMenuGOptionAppartientGroupe (NumGroupe, NumTag : integer) : boolean;
  begin
    { GC_DBR_20081023_FQ;034;15656 }
    Result := (NumGroupe = -1) or bNumModule or FMenuG.OptionAppartientGroupe (NumGroupe, NumTag);
  end;

BEGIN
Case VH_GC.NumModuleCourant of
{$IF Defined(CRM) or Defined(GRCLIGHT)}
   65 :  BEGIN { Gestion des paramètres}
{$ELSE}
   92 :  BEGIN { Gestion des paramètres}
{$IFEND}
        if FMenuGOptionAppartientGroupe (NumGroupe, -92901) then
        begin
         for i:=0 to 9 do
            begin
            St:=RechDom('RTLIBCHAMPSLIBRES','ML'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92902+i,St) else FMenuG.RemoveItem(92902+i);
            end;
              //mcd 20/04/2007 pour nouvelle tabletet GRC
         St:=RechDom('RTLIBCHAMPSLIBRES','MLA',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92690,St) else FMenuG.RemoveItem(92690);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLB',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92691,St) else FMenuG.RemoveItem(92691);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLC',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92692,St) else FMenuG.RemoveItem(92692);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLD',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92693,St) else FMenuG.RemoveItem(92693);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLE',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92694,St) else FMenuG.RemoveItem(92694);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLG',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92696,St) else FMenuG.RemoveItem(92696);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLH',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92697,St) else FMenuG.RemoveItem(92697);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLI',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92698,St) else FMenuG.RemoveItem(92698);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLJ',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92699,St) else FMenuG.RemoveItem(92699);
         St:=RechDom('RTLIBCHAMPSLIBRES','MLF',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92695,St) else FMenuG.RemoveItem(92695);
        end;
          //fin mcd 20/04/2007
        if FMenuGOptionAppartientGroupe (NumGroupe, 92951) then
        begin
         for i:=1 to 35 do
         // Tables libres
            begin
            if i < 11 then St:=RechDom('RTLIBCHAMPSLIBRES','CL'+IntToStr(i-1),FALSE)
                      else St:=RechDom('RTLIBCHAMPSLIBRES','CL'+Chr(i-1+55),FALSE) ;
            if copy(St,1,1) = '.' then FMenuG.RemoveItem(92950+i) else FMenuG.RenameItem(92950+i,St)
            end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92990) then
        begin
         for i:=1 to 3 do
            begin
            // Actions
            St:= RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92990+i,St) else FMenuG.RemoveItem(92990+i);
            // Perspectives
            St:= RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92995+i,St) else FMenuG.RemoveItem(92995+i);
            end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92931) then
        begin
         for i:=1 to 10 do
            begin
            //RTLIBTABLECOMPL
            St:= RechDom('RTLIBTABLECOMPL','TL'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92931+i,St) else FMenuG.RemoveItem(92931+i);
            end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92919) then
        begin
         // Suspect
         for i:=0 to 9 do
            begin
            St:=RechDom('RTLIBCHPLIBSUSPECTS','CL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92920+i,St) else FMenuG.RemoveItem(92920+i);
            end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92611) then
        begin
          // ciblage
         for i:=1 to 5 do
            begin
            St:=RechDom('RTLIBCIBLAGE','TL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92611+i,St) else FMenuG.RemoveItem(92611+i);
            end;
          // projet : mng 24/05/07 uniquement si gestion des projets
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92621) then
        begin
         if GetParamsocSecur('SO_RTPROJGESTION', false) then
           for i:=1 to 5 do
              begin
              St:=RechDom('RTLIBPROJET','TL'+IntToStr(i),FALSE);
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92621+i,St) else FMenuG.RemoveItem(92621+i);
              end;
              for i:= 6 to 9 do
              begin
                St:=RechDom('RTLIBPROJET','TL'+IntToStr(i),FALSE);
                if copy(St,1,1) <> '.' then FMenuG.RenameItem(92521+i-5,St) else FMenuG.RemoveItem(92521+i-5);
              end;
              St:=RechDom('RTLIBPROJET','TLA',FALSE);
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92526,St) else FMenuG.RemoveItem(92526);
           // Propositions marketing
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92631) then
        begin
         for i:=1 to 5 do
            begin
            St:=RechDom('RTLIBPERSPECTIVE','TL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92631+i,St) else FMenuG.RemoveItem(92631+i);
            end;
          for i:=6 to 9 do
            begin
            St:=RechDom('RTLIBPERSPECTIVE','TL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92531+i-5,St) else FMenuG.RemoveItem(92531+i-5);
            end;
          St:=RechDom('RTLIBPERSPECTIVE','TLA',FALSE);
          if copy(St,1,1) <> '.' then FMenuG.RenameItem(92536,St) else FMenuG.RemoveItem(92536);
          // Opérations
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92641) then
        begin
         for i:=1 to 5 do
            begin
            St:=RechDom('RTLIBOPERATION','TL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92641+i,St) else FMenuG.RemoveItem(92641+i);
            end;
         for i:=6 to 9 do
            begin
            St:=RechDom('RTLIBOPERATION','TL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92541+i-5,St) else FMenuG.RemoveItem(92541+i-5);
            end;
         St:=RechDom('RTLIBOPERATION','TLA',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92546,St) else FMenuG.RemoveItem(92546);
        end;
             // actions 3jl le 26/06/2008
        
        if FMenuGOptionAppartientGroupe (NumGroupe, -92550) then
        begin
         for i:=1 to 9 do
            begin
            St:=RechDom('RTLIBACTIONS','TL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92558+i,St) else FMenuG.RemoveItem(92558+i);
            end;
         St:=RechDom('RTLIBACTIONS','TLA',FALSE);
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92568,St) else FMenuG.RemoveItem(92568);
        end;
         { Libellé des infos complémentaires Contacts mng 24/05/07 uniquement si gestion des ..}
         if GetParamsocSecur('SO_RTGESTINFOS006', false) then
          if FMenuGOptionAppartientGroupe (NumGroupe, -92785) then
          begin
            for i:=1 to 5 do
                begin
                // tables
                St:= RechDom('RTLIBCHAMPS006','CL'+IntToStr(i-1),FALSE) ;
                if copy(St,1,1) <> '.' then FMenuG.RenameItem(92785+i,St) else FMenuG.RemoveItem(92785+i);
                // multi-choix
                St:= RechDom('RTLIBCHAMPS006','ML'+IntToStr(i-1),FALSE) ;
                if copy(St,1,1) <> '.' then FMenuG.RenameItem(92792+i,St) else FMenuG.RemoveItem(92792+i);
                end;
          end;

            //RTLIBTABLECOMPL table 11 : A
          if FMenuGOptionAppartientGroupe (NumGroupe, -92641) then
          begin
            St:= RechDom('RTLIBTABLECOMPL','TLA',FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92942,St) else FMenuG.RemoveItem(92942);
            //RTLIBTABLECOMPL table 11 : A
            St:= RechDom('RTLIBTABLECOMPL','TLB',FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92943,St) else FMenuG.RemoveItem(92943);
            if not GetParamsocSecur('SO_RTCONFIDENTIALITE', false) then FMenuG.RemoveItem(92946) ;
          end;
{$IFDEF CTI}
            if not GetParamsocSecur('SO_RTCTIGESTION', false) then
            begin
              FMenuGRemoveItem(92760) ;
            end;
{$ELSE}
            FMenuGRemoveItem(92760) ;
{$ENDIF}
        if FMenuGOptionAppartientGroupe (NumGroupe, -92700) then
        begin        
         // tables libres et multi-choix infos compl. actions mng 24/05/07 uniquement si gestion des ..
         if GetParamsocSecur('SO_RTGESTINFOS001', false) then
           for i:=1 to 5 do
              begin
              // tables
              St:= RechDom('RTLIBCHAMPS001','CL'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92700+i,St) else FMenuG.RemoveItem(92700+i);
              // multi-choix
              St:= RechDom('RTLIBCHAMPS001','ML'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92720+i,St) else FMenuG.RemoveItem(92720+i);
              end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92710) then
        begin
         // tables libres et multi-choix infos compl. operations mng 24/05/07 uniquement si gestion des ..
         if GetParamsocSecur('SO_RTGESTINFOS002', false) then
           for i:=1 to 5 do
              begin
              // tables
              St:= RechDom('RTLIBCHAMPS002','CL'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92710+i,St) else FMenuG.RemoveItem(92710+i);
              // multi-choix
              St:= RechDom('RTLIBCHAMPS002','ML'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92730+i,St) else FMenuG.RemoveItem(92730+i);
              end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92780) then
        begin
         // Chainages d'actions
         for i:=1 to 3 do
            begin
            St:= RechDom('RTLIBCHAINAGES','RH'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92780+i,St) else FMenuG.RemoveItem(92780+i);
            end;
        end;
        // au menu juste pour appel depuis menu contextuel GRC

        // Pl le 14/08/07 : FQ10622 on veut tous les nouveaux champs libre maintenant
(*        if (GetParamSocSecur ('SO_AFCLIENT', 0) <>8 ) then  //cInClientKPMG =8 pour ne pas inclure menu spécif
           begin //mcd 20/04/07
          if FMenuGOptionAppartientGroupe (NumGroupe, 92630) then
          begin
           FMenuG.RemoveItem(92630); //lib table libre prospect sur 6c
           FMenuG.RemoveItem(-92631); //table libre prospect sur 6c
          end;
          if FMenuGOptionAppartientGroupe (NumGroupe, 92630) then
          begin
           FMenuG.RemoveItem(92640); //lib table libre opération
           FMenuG.RemoveItem(-92641); //table libre opération
          end;
           end;*)
        // Pl le 14/08/07 : FQ10622 on veut tous les nouveaux champs libre maintenant
//         if (GetParamSocSecur ('SO_AFCLIENT', 0) <>8 ) or (not GetParamsocSecur('SO_RTPROJGESTION', false)) then
         if (not GetParamsocSecur('SO_RTPROJGESTION', false)) then
           begin
           FMenuGRemoveItem(92620); //table libre projet
           FMenuGRemoveItem(-92621); //table libre projet
           FMenuG.RemoveItem(-92627); //libelle projet 10931
           FMenuG.RemoveItem(92671); //type projet 10931
           end;
         if not GetParamsocSecur('SO_RTGESTINFOS00Q', false) then
           begin
           FMenuGRemoveItem(92670); //info compl projet
           FMenuGRemoveItem(-92651); //table info compl projet
           FMenuGRemoveItem(-92652); //table muilti choix projet
           end;
         if not GetParamsocSecur('SO_RTGESTINFOS00V', false) then
           begin
           FMenuGRemoveItem(92865); //info compl proposition
           FMenuGRemoveItem(-92870); //table info compl proposition
           FMenuGRemoveItem(-92880); //table muilti choix proposition
           end;

        if FMenuGOptionAppartientGroupe (NumGroupe, -92651) then
        begin
         // tables libres et multi-choix infos compl. projets  mng 24/05/07 uniquement si gestion des ..
         if GetParamsocSecur('SO_RTGESTINFOS00Q', false) then
           for i:=1 to 5 do
              begin
              // tables
              St:= RechDom('RTLIBCHAMPS00Q','CL'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92652+i,St) else FMenuG.RemoveItem(92652+i);
              // multi-choix
              St:= RechDom('RTLIBCHAMPS00Q','ML'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92657+i,St) else FMenuG.RemoveItem(92657+i);
              end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, -92870) then
        begin
         // tables libres et multi-choix infos compl. projets  mng 24/05/07 uniquement si gestion des ..
         if GetParamsocSecur('SO_RTGESTINFOS00V', false) then
           for i:=1 to 5 do
              begin
              // tables
              St:= RechDom('RTLIBCHAMPS00V','CL'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92870+i,St) else FMenuG.RemoveItem(92870+i);
              // multi-choix
              St:= RechDom('RTLIBCHAMPS00V','ML'+IntToStr(i-1),FALSE) ;
              if copy(St,1,1) <> '.' then FMenuG.RenameItem(92880+i,St) else FMenuG.RemoveItem(92880+i);
              end;
        end;
{$IFDEF GIGI}
        if not GetParamsocSecur('SO_RTSUIVIAVANCEMENT', False) then
         Fmenug.removeItem(92948);
        if not GetParamsocSecur('SO_RTGESTPARTENAIRES', False) then
         begin
//         Fmenug.removeItem(92390); // suivi des frais mcd 30/07/08
         Fmenug.removeItem(92380); // AG - 17.07.08
         end;
        if not GetParamsocSecur('SO_AFRTOPERATIONS', False) then
         Fmenug.removeItem(92390); // AG - 06.08.08 (Demande PL)
        //Gestion accès fonction avancées SUSPECTS
        if not GetParamSocSecur('SO_AFRTSUSPECT', False, True) then
        begin
          FMenuGRemoveItem(92135); //suppresion suspect
          FMenuGRemoveItem(92262); //bascule suspects vers prospects
          FMenuGRemoveItem(92999); //table de correspondances
          FMenuGRemoveItem(92266); //recherche mixte
          FMenuGRemoveItem(92990); //RAZ correspondance
        end;
        //Gestion accès fonction avancées OPERATIONS
        if not GetParamSocSecur('SO_AFRTOPERATIONS', False, True) then
        begin
          FMenuGRemoveItem(-92260);  //Ciblage
          FMenuGRemoveItem(-92253);  //Suivi opérations
        end;
{$ENDIF GIGI}
        if not GetParamsocSecur('SO_RTGESTINFOS001', false) then
        begin
          FMenuGRemoveItem(-92700) ;
          FMenuGRemoveItem(-92720) ;
          FMenuGRemoveItem(92740) ;
        end;

        if not GetParamsocSecur('SO_RTGESTINFOS002', false) then
        begin
          FMenuGRemoveItem(-92710) ;
          FMenuGRemoveItem(-92730) ;
          FMenuGRemoveItem(92770) ;
        end;

        if not GetParamsocSecur('SO_RTGESTINFOS006', false) then
        begin
          FMenuGRemoveItem(92287) ;
          FMenuGRemoveItem(-92785) ;
          FMenuGRemoveItem(-92792) ;
          FMenuGRemoveItem(92765) ;
        end;


{$ifndef GIGI}
         FMenuGRemoveItem(92944) ;  // saisi info complémentaire DP pour la GA
{$else}
        if not GetParamsocSecur('SO_AFLIENDP', '-') then FMenuGRemoveItem(92944) ;  // mcd 10/05/06 inutile si pa slien DP
{$ENDIF GIGI}
        if V_PGI.NumVersionSoc > 700 then
           if not GetParamsocSecur('SO_RTGESTIONGED', false) then
           begin
             FMenuGRemoveItem(92681);
             FMenuGRemoveItem(-92682);
           end
        else
        if FMenuGOptionAppartientGroupe (NumGroupe, -92682) then
        begin
          for i:=1 to 3 do ChangeTitreOuVireMenu(92682+i,'RTLIBGED','RD'+IntToStr(i)) ;
        end;
         { Multi choix Base de connaissance 92831 }
        if FMenuGOptionAppartientGroupe (NumGroupe, 92831) then
        begin
         St:= RechDom('RBMULTILIBREBC','BM1',FALSE) ;
         if copy(St,1,1) <> '.' then FMenuG.RenameItem(92831,St) else FMenuG.RemoveItem(92831);
        end;
        END ;
//      End ; (* Case *)

//Case NumModule of
   93 :  BEGIN { Champs libres Combos Prospect}
         if not GetParamsocSecur('SO_RTGESTINFOS003', False) then
         begin
           FMenuG.RemoveItem(93510) ;
           FMenuG.RemoveItem(-93520) ;
           FMenuG.RemoveItem(-93540) ;
         end;
        // tables libres et multi-choix infos compl. fournisseurs
        if FMenuGOptionAppartientGroupe (NumGroupe, 93521) then
        begin
         for i:=1 to 5 do
            begin
            // tables
            St:= RechDom('RTLIBCHAMPS003','CL'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(93520+i,St) else FMenuG.RemoveItem(93520+i);
            // multi-choix
            St:= RechDom('RTLIBCHAMPS003','ML'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(93540+i,St) else FMenuG.RemoveItem(93540+i);
            end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, 93631) then
        begin
         { libres actions fournisseurs }
         for i:=1 to 3 do
            begin
            // Actions
            St:= RechDom('RFLIBCLACTIONS','FA'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(93630+i,St) else FMenuG.RemoveItem(93630+i);
            end;
        end;
        if FMenuGOptionAppartientGroupe (NumGroupe, 93671) then
        begin
         { libres chainages fournisseurs }
         for i:=1 to 3 do
            begin
            // Actions
            St:= RechDom('RFLIBCHAINAGES','RF'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(93670+i,St) else FMenuG.RemoveItem(93670+i);
            end;
        end;
         if not GetParamsocSecur('SO_RFCONFIDENTIALITE', false) then FMenuGRemoveItem(93570) ;
         if not GetParamsocSecur('SO_RTGESTINFOS003', false) then
            FMenuGRemoveItem(93130) ;
         END ;
      End ; (* Case *)
   Case VH_GC.NumModuleCourant of
   92,311 :  BEGIN { Gestion des Items des traitements}
{$IF not defined(GRCLIGHT) or defined(CRM) }
            if not GetParamsocSecur('SO_RTPROJGESTION', false) then
{$IFEND}
              FMenuGRemoveItem(92105) ;
{$IFDEF CTI}
            if not GetParamsocSecur('SO_RTCTIGESTION', false) then
            begin
              FMenuGRemoveItem(92205) ;
            end;
{$ELSE}
            FMenuGRemoveItem(92205) ;
{$ENDIF}
        if not GetParamsocSecur('SO_RTGESTINFOS001', false) then
        begin
          FMenuGRemoveItem(92255) ;
          FMenuGRemoveItem(92256) ;
        end;
        if not GetParamsocSecur('SO_RTGESTINFOS00V', false) then { modif série infos compl propal }
          FMenuGRemoveItem(92257) ;
        if not GetParamsocSecur('SO_RTGESTINFOS006', false) then
          FMenuGRemoveItem(92287) ;
        // au menu juste pour appel depuis menu contextuel GRC
//        FMenuGRemoveItem(92296); // Fiche "Mes Actions"
        // transfert suspects vers prospects
        if (not ExJaiLeDroitConcept(TConcept(GcTransfertSusPro), False)) then
          FMenuGRemoveItem(92262);
{$IFDEF AFFAIRE}
        if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) then
{$ENDIF}
           begin
           FMenuGRemoveItem(92226); // Mailing Affaires
           FMenuGRemoveItem(92216); // Export Mailing Affaires
           FMenuGRemoveItem(92236); // Generations actions Affaires
           FMenuGRemoveItem(92418); // analyse dynamique Affaires
           FMenuGRenameItem(92440, TraduireMemoire('Propositions/Devis'));
           FMenuGRemoveItem(27090); // menu pop affaire
           end;
{$IFDEF AFFAIRE}
        if ( VH_GC.GASeria=false ) then
          begin
          FMenuGRenameItem(92440, TraduireMemoire('Propositions/Devis'));
          FMenuGRemoveItem(92418); // analyse dynamique Affaires
          end;
{$ENDIF}
         if not VH_GC.GCIfDefCEGID then
         begin
           FMenuGRemoveItem(92403);
           //FMenuG.RemoveItem(92440);
         end;
        if V_PGI.NumVersionSoc > 700 then
           if not GetParamsocSecur('SO_RTGESTIONGED', false) then
           begin
             FMenuGRemoveItem(92112);
           end;
{$IFDEF GRCLIGHT}
        { fichiers }
        FMenuGRemoveItem(92103);
        FMenuGRemoveItem(92106); FMenuGRemoveItem(92113);
        FMenuGRemoveItem(92135); //suppresion suspect
        FMenuGRemoveItem(92266); //recherche client/prospect/suspect

        Fmenug.removeItem(92380); // AG - 17.07.08 // suivi des Partenaires
        { Traitement}
        FMenuGRemoveItem(-92260);  //Ciblage
        FMenuGRemoveItem(-92230); FMenuGRemoveItem(-92220);FMenuGRemoveItem(-92210);
        FMenuGRemoveItem(92288); FMenuGRemoveItem(92291); FMenuGRemoveItem(92270);
        FMenuGRemoveItem(92262); //bascule suspects vers prospects
        { éditions}
        FMenuGRemoveItem(-92320); FMenuGRemoveItem(92370);
        { analyses } FMenuGRemoveItem(92419); FMenuGRemoveItem(92425); FMenuGRemoveItem(92426);
        { mng FQ 10771 }
        if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
          FMenuGRemoveItem(92460);
{$ENDIF GRCLIGHT}
         END ;
      End ; (* Case *)
{$IFDEF GRCLIGHT}
   Case VH_GC.NumModuleCourant of
   65 :  BEGIN { Paramètres CRM}
{$ifndef GIGI}
         FMenuGRemoveItem(92944) ;  // saisi info complémentaire DP pour la GA
{$else}
         if not GetParamsocSecur('SO_AFLIENDP', '-') then FMenuGRemoveItem(92944) ;  // mcd 10/05/06 inutile si pa slien DP
{$ENDIF GIGI}
         if (not GetParamsocSecur('SO_RTPROJGESTION', false)) then
           begin
           FMenuG.RemoveItem(-92627); //libelle projet 10931
           FMenuG.RemoveItem(92671); //type projet 10931
           end;
         FMenuGRemoveItem(-92810); FMenuGRemoveItem(-92830); FMenuGRemoveItem(92913); FMenuGRemoveItem(-92919);
         FMenuGRemoveItem(92914); FMenuGRemoveItem(92947);FMenuGRemoveItem(92609);
         FMenuGRemoveItem(92770); FMenuGRemoveItem(-92710); FMenuGRemoveItem(-92730); FMenuGRemoveItem(92916);
         FMenuGRemoveItem(-92611); FMenuGRemoveItem(92620);  FMenuGRemoveItem(-92621);
         FMenuGRemoveItem(92650); FMenuG.RemoveItem(92670);
         FMenuGRemoveItem(-92651); FMenuGRemoveItem(-92652); FMenuGRemoveItem(92640);
         FMenuGRemoveItem(-92641); FMenuGRemoveItem(92999); FMenuGRemoveItem(92990);
         FMenuGRemoveItem(-92627);
         FMenuGRemoveItem(-92610);
         END ;
      End ; (* Case *)
{$ENDIF GRCLIGHT}
END ;

{$IFDEF ANCIENNEVERSION}
Procedure TraiteMenusGRC ( NumModule : integer ) ;
Var i : integer ;
    St : String ;
BEGIN
Case NumModule of
   92 :  BEGIN { Champs libres Combos Prospect}
         for i:=0 to 9 do
            begin
            St:=RechDom('RTLIBCHAMPSLIBRES','ML'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92980+i,St) else FMenuG.RemoveItem(92980+i);
            end;
         for i:=1 to 26 do
         // Tables libres
            begin
            if i < 11 then St:=RechDom('RTLIBCHAMPSLIBRES','CL'+IntToStr(i-1),FALSE)
                      else St:=RechDom('RTLIBCHAMPSLIBRES','CL'+Chr(i-1+55),FALSE) ;
            if copy(St,1,1) = '.' then FMenuG.RemoveItem(92950+i) else FMenuG.RenameItem(92950+i,St)
            end;
         for i:=1 to 3 do
            begin
            // Actions
            St:= RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92990+i,St) else FMenuG.RemoveItem(92990+i);
            // Perspectives
            St:= RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92995+i,St) else FMenuG.RemoveItem(92995+i);
            end;
         for i:=1 to 10 do
            begin
            //RTLIBTABLECOMPL
            St:= RechDom('RTLIBTABLECOMPL','TL'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92931+i,St) else FMenuG.RemoveItem(92931+i);
            end;
         // Suspect
         for i:=0 to 9 do
            begin
            St:=RechDom('RTLIBCHPLIBSUSPECTS','CL'+IntToStr(i),FALSE);
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92920+i,St) else FMenuG.RemoveItem(92920+i);
            end;

            //RTLIBTABLECOMPL table 11 : A
            St:= RechDom('RTLIBTABLECOMPL','TLA',FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92942,St) else FMenuG.RemoveItem(92942);
            //RTLIBTABLECOMPL table 11 : A
            St:= RechDom('RTLIBTABLECOMPL','TLB',FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92943,St) else FMenuG.RemoveItem(92943);
            if GetParamSoc('SO_RTPROJGESTION') = False then FMenuG.RemoveItem(92105) ;
            if GetParamSoc('SO_RTCONFIDENTIALITE') = False then FMenuG.RemoveItem(92946) ;
{$IFDEF CTI}
            if GetParamSoc('SO_RTCTIGESTION') = False then FMenuG.RemoveItem(92760) ;
            if GetParamSoc('SO_RTCTIGESTION') = False then FMenuG.RemoveItem(92205) ;
{$ELSE}
            FMenuG.RemoveItem(92205) ;
            FMenuG.RemoveItem(92760) ;
{$ENDIF}
         // tables libres et multi-choix infos compl. actions
         for i:=1 to 5 do
            begin
            // tables
            St:= RechDom('RTLIBCHAMPS001','CL'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92700+i,St) else FMenuG.RemoveItem(92700+i);
            // multi-choix
            St:= RechDom('RTLIBCHAMPS001','ML'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92720+i,St) else FMenuG.RemoveItem(92720+i);
            end;
         // tables libres et multi-choix infos compl. operations
         for i:=1 to 5 do
            begin
            // tables
            St:= RechDom('RTLIBCHAMPS002','CL'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92710+i,St) else FMenuG.RemoveItem(92710+i);
            // multi-choix
            St:= RechDom('RTLIBCHAMPS002','ML'+IntToStr(i-1),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92730+i,St) else FMenuG.RemoveItem(92730+i);
            end;
         // Chainages d'actions
         for i:=1 to 3 do
            begin
            St:= RechDom('RTLIBCHAINAGES','RH'+IntToStr(i),FALSE) ;
            if copy(St,1,1) <> '.' then FMenuG.RenameItem(92780+i,St) else FMenuG.RemoveItem(92780+i);
            end;

         {$IFDEF EAGLCLIENT}
         //FMenuG.RemoveItem(92201);// télévente // mng 04-01-02 FMenuG.RemoveItem(92202); // FMenuG.RemoveItem(92203) {Mailing};
         //FMenuG.RemoveItem(92270);FMenuG.RemoveItem(92290); // import suspect et dédoublonnage
         FMenuG.RemoveGroup(92800,True);   //Outils
         //FMenuG.RemoveGroup(92900,True);   //Parametres généraux
         //FMenuG.RemoveGroup(92500,True);   //Parametres actions/propositions
//         FMenuG.RemoveItem(92111);   //Suppression tiers
         {$ENDIF}
        // au menu juste pour appel depuis menu contextuel GRC
        FMenuG.RemoveItem(92296); // Fiche "Mes Actions"
        FMenuG.RemoveItem(92948); // tablette libellés tables libres table complémentaire

        if GetParamSoc('SO_RTGESTINFOS001') = False then
            begin
            FMenuG.RemoveItem(-92700) ;
            FMenuG.RemoveItem(-92720) ;
            FMenuG.RemoveItem(92740) ;
            end;

        if GetParamSoc('SO_RTGESTINFOS002') = False then
            begin
            FMenuG.RemoveItem(-92710) ;
            FMenuG.RemoveItem(-92730) ;
            FMenuG.RemoveItem(92770) ;
            end;

        if (GetParamSoc('SO_RTPROPTABHIE') = False) and (GetParamSoc('SO_RTACTTABHIE') = False) then
            FMenuG.RemoveItem(92949) ;


        if Not VH_GC.GASeria then
           begin
           FMenuG.RemoveItem(92226); // Mailing Affaires
           FMenuG.RemoveItem(92216); // Export Mailing Affaires
           FMenuG.RemoveItem(92236); // Generations actions Affaires
           end;

{$IFDEF CCS3}
        FMenuG.RemoveItem(92107);  // propositions
        FMenuG.RemoveItem(92106);  // suspects
        FMenuG.RemoveItem(92108);  // concurrents
        FMenuG.RemoveItem(92201);  // téléventes
        FMenuG.RemoveItem(92270);  // import de suspects
        FMenuG.RemoveItem(92271);  // mailing suspect
        FMenuG.RemoveItem(92272);  // export suspect
        FMenuG.RemoveItem(92291);  // dédoublonnage
        FMenuG.RemoveItem(92288);  // qualité suspect
        FMenuG.RemoveItem(-92340);  // édition propositions
        FMenuG.RemoveItem(92413); // analyse dynamique propositions
        FMenuG.RemoveItem(92414); // analyse concurrence
        FMenuG.RemoveItem(92417); // analyse pièces/propositions
        FMenuG.RemoveItem(92423); // analyse croisées propositions
        FMenuG.RemoveItem(92424); // analyse croisées détail propositions
        FMenuG.RemoveItem(92425); // analyse croisées suspects
        FMenuG.RemoveItem(92440); // Propositions sans devis
        FMenuG.RemoveItem(92912); // tablettes types de propositions
        FMenuG.RemoveItem(92917); // Motifs de signature / perte des propositions
        FMenuG.RemoveItem(92913);  // Libellés champs libres Suspects
        FMenuG.RemoveItem(92914);  // Motifs fermeture  fiche Suspect
        FMenuG.RemoveItem(-92995); // Tables libres propositions (...96 à 98 )
        FMenuG.RemoveItem(92946); // confidentialité
        FMenuG.RemoveItem(92947); // Correspondance Suspect/Prospect
        FMenuG.RemoveItem(-92919); // tables libres suspects (...20 à 29 )
        FMenuG.RenameItem(92915,'Libellés tables libres actions');
        //FMenuG.RenameGroup(92500,'Paramètres Actions',True);
        FMenuG.RemoveItem(-92700) ;
        FMenuG.RemoveItem(-92720) ;
        FMenuG.RemoveItem(92740) ;
        FMenuG.RemoveItem(92750) ;
        FMenuG.RemoveItem(92945) ;  // paramétrage table complémentaire
        FMenuG.RemoveItem(-92931) ;  // tables libres table complémentaire
        FMenuG.RemoveItem(92948) ;  // libellés table libre complémentaire
        FMenuG.RemoveGroup(92600, True);   // Tarif fournisseurs
{$ENDIF}
         if not VH_GC.GCIfDefCEGID then
         begin
           FMenuG.RemoveItem(92403);
           //FMenuG.RemoveItem(92440);
           FMenuG.RemoveItem(92850);
{$IFNDEF CTI}
           FMenuG.RemoveItem(92880);
           FMenuG.RemoveItem(92890);
{$ENDIF}
         end;

         END ;
      End ; (* Case *)
END ;
{$ENDIF}

Procedure TripoteMenuGCS3 ( NumModule : integer ) ;

          Procedure RemoveTablesLibresS3 ( TT : Array of integer ) ;
          Var i : integer ;
          BEGIN
          for i:=Low(TT) to High(TT) do FMenuG.RemoveItem(TT[i]) ;
          END ;

BEGIN
{$IFDEF CCS3}
Case NumModule of
   30 : BEGIN
        FMenuG.RemoveItem(30237) ; {Duplication Facture provisoire}
//        FMenuG.RemoveItem(30401) ; {Visa des pièces} // DBR - fiche 10865
        FMenuG.RemoveItem(30800) ; {Edition affaire}
        END ;
   31 : BEGIN
//        FMenuG.RemoveItem(31501) ; {Visa des pièces} // DBR - Fiche 10865
        END ;
   32 : BEGIN
        //FMenuG.RemoveItem(32505) ;       {Traductions articles}
        FMenuG.RemoveItem(32602) ;       {Emplacements}
        FMenuG.RemoveGroup(32700,True) ; {Dimensions}
        END ;
   33 : BEGIN
        {$IFDEF PREMIUM}
         if not VH_GC.GCAchatStockSeria then FMenuG.RemoveGroup(33200,True); // Achats
        {$ENDIF}
        FMenuG.RemoveItem(-33130) ; {Analyses par affaires}
        END ;
   60 : BEGIN
        {$IFDEF PREMIUM}
          if not VH_GC.GCAchatStockSeria then
          begin
            FMenuG.RemoveItem(60402); // Vérif des stocks
            FMenuG.RemoveItem(-30440); // Affectation de commandes
            FMenuG.RemoveItem(60412) ; {Historisation des pièces Achat}
            FMenuG.RemoveItem(60422) ; {Epuration des pièces Achat}
          end;
        {$ENDIF}
        END;
   65 : BEGIN
        FMenuG.RemoveItem(65102)  ; {Mode expédition}
        FMenuG.RemoveItem(65204)  ; {Assistant code affaire}
        FMenuG.RemoveItem(65305)  ; {Types emplacement}
        FMenuG.RemoveItem(65306)  ; {Cotes emplacements}
        FMenuG.RemoveItem(65410)  ; {Domaines activité}
        FMenuG.RemoveItem(65406)  ; {Emploi mémos/photos}
        FMenuG.RemoveItem(65473)  ; {Listes de saisie}
        FMenuG.RemoveItem(65475)  ; {Natures de regoupement}
        FMenuG.RemoveItem(-65420) ; {Affaires}
        FMenuG.RemoveItem(-65520) ; {Libres affaires}
        FMenuG.RemoveItem(-65524) ; {Lib Libres affaires}
        FMenuG.RemoveItem(-65250) ; {Champs libres Affaires}
        FMenuG.RemoveItem(-65530) ; {Libres ressources}
        FMenuG.RemoveItem(-65534) ; {Lib Libres ressources}
        FMenuG.RemoveItem(-65260) ; {Champs libres Ressources}
        FMenuG.RemoveItem(-65540) ; {Pièces}
        FMenuG.RemoveItem(-65550) ; {Etablissements}
        FMenuG.RemoveItem(-65580) ; {Libellés Etab.}
        FMenuG.RemoveItem(65924)  ; {Langues}
        FMenuG.RemoveItem(65923)  ; {Jours fériés}
        FMenuG.RemoveGroup(65600,True) ;
        // grc
        //FMenuG.RemoveGroup(65650,True) ;{Tablettes affaire, ressources}
        FMenuG.RemoveItem(65652)  ; {Codes tarifs}
        FMenuG.RemoveItem(65654)  ; {Compétences}
        FMenuG.RemoveItem(65655)  ; {Niveau de diplome}
        FMenuG.RemoveItem(65656)  ; {Type de calendrier}

        RemoveTablesLibresS3([65562,65563,65564,65565,65504,65505,65506,65507,65508,65509,65510]) ;
        RemoveTablesLibresS3([65572,65573,65574,65575,65514,65515,65516,65517,65518,65519,65520]) ;
        RemoveTablesLibresS3([65596,65597]) ;
        RemoveTablesLibresS3([65587,65588,65589,65590]) ;
        RemoveTablesLibresS3([65587,65588,65589,65590]) ;
        RemoveTablesLibresS3([65582,65583,65584,65585,65555,65556,65557,65558,65559,65560]) ;
        RemoveTablesLibresS3([65674,65675,65676,65677,65678,65679,65680]) ;
        END ;
   END ;
{$ENDIF}
END ;


Procedure TraiteSpecifsCEGID ( NumModule : integer ) ;
BEGIN
if VH_GC.GCIfDefCEGID then
begin
   Case NumModule of
   33 : BEGIN
        FMenuG.RemoveItem(33201);     // Analyse Derniers prix d'achats
        END ;
   END ;
end
else
begin
Case NumModule of
   30 : begin
        FMenuG.RemoveGroup(30700,True) ;  {Suivi Client}
        FMenuG.RemoveItem(30610); FMenuG.RemoveItem(30611); //portefeuilles spécifiques CEGID
        FMenuG.RemoveItem(-30620) ; FMenuG.RemoveItem(-30640) ; {Specifs Cegid}
        FMenuG.RemoveItem(30609) ; // Edition avec code mise sous pli
        end;
   33 : begin
        FMenuG.RemoveItem(33106);  //stats Cegid (à mettre seulement dans IFNDEF CEGID quand operationnel
        FMenuG.RemoveItem(33107);  // Nouveaux clients
        FMenuG.RemoveItem(-33140) ; FMenuG.RemoveItem(-33150) ; // stats specif Cegid
        FMenuG.RemoveItem(-33160) ; FMenuG.RemoveItem(-33170) ; // stats specif Cegid
        FMenuG.RemoveItem(33106);  //stats Cegid
        FMenuG.RemoveItem(-33180);  //Stat commission IC
        FMenuG.RemoveItem(33181);  //Stat commission IC
        FMenuG.RemoveItem(33182);  //Cube commission IC
        end;
   92 : begin
        FMenuG.RemoveItem(92403);  // Activité par commercial
        end;
   end;
end;
END ;

Function TraiteMenuSpecif ( NumModule : integer ) : boolean;
var stArguments,stType,stInfos : string;
    Domaine,NomFiche,Range,Lequel,Argument : string;
begin
Result := true;
stArguments := RechDom('YYMENUSPECIF',IntToStr(NumModule),False);
stType := RechDom('YYMENUSPECIF',IntToStr(NumModule),True);
if stType = 'FIC' then
   begin
   Domaine := ReadTokenSt(stArguments); NomFiche := ReadTokenSt(stArguments);
   Range  :=ReadTokenSt(stArguments); Lequel :=ReadTokenSt(stArguments);
   Argument := ReadTokenSt(stArguments);
   AGLLanceFiche(Domaine,NomFiche,Range,Lequel,Argument);
   end else
   if stType = 'EXE' then
      begin
      stInfos := ' /USER='+V_PGI.UserLogin+' /PASSWORD='+V_PGI.PassWord
          +' /DATE='+DateToStr(V_PGI.DateEntree)+' /DOSSIER='+v_pgi.CurrentAlias+' /MAJSTRUCTURE=TRUE';
      WinExec(PChar(stArguments + stInfos), sw_Show);
      end else Result := false;
end;

// Cette procédure permet d'ajuster chaque module du menu, en fonction de la Séria
procedure AjusteMenu ( NumModule : integer ) ;
BEGIN
  Case NumModule of
  65  : ;
  102 : if (CtxMode in V_PGI.PGIContexte) then  // MODE
        begin   // Ventes
        if (V_PGI.VersionDemo = False) then
           begin
           if (VH_GC.BOSeria = False) then
              begin
              FMenuG.RemoveGroup(102100,True);  // Saisie des Ventes Détail
              FMenuG.RemoveGroup(30900,True);   // Tarif Détail
              end;
           if (VH_GC.GCSeria = False) then
              begin
              FMenuG.RemoveGroup(30200,True);   // Saisie des Ventes Négoce
              FMenuG.RemoveGroup(30400,True);   // Génération documents
              FMenuG.RemoveGroup(30100,True);   // Tarif Négoce
              end;
           end;
        end ;
  end;
END;

Procedure DispatchTTArticle ( Action : TActionFiche ; Lequel,TT,Range : String ) ;
var x : integer ;
    TypeArticle,Critere,Valeur,Champ,ArticleSpecif : string ;
begin
if ((Lequel='') and (Action=taCreat)) then
   BEGIN
   TypeArticle:='MAR' ;
   Critere:=(ReadTokenSt(Range));
   While (Critere <>'') do
      BEGIN
      X:=pos('=',Critere);
      if x<>0 then
         begin
         Champ:=copy(Critere,1,X-1);
         Valeur:=Copy (Critere,X+1,length(Critere)-X);
         end;
      if Champ = 'TYPEARTICLE' then TypeArticle := Valeur;
      Critere:=(Trim(ReadTokenSt(Range)));
      END;
   END else
   BEGIN
   if Not IsCodeArticleUnique(Lequel) then Lequel:=CodeArticleUnique(Lequel,'','','','','') ;
   TypeArticle:=GetChampsArticle(Lequel,'GA_TYPEARTICLE') ;
   END ;
ArticleSpecif:=IsArticleSpecif(TypeArticle) ;
if ArticleSpecif='FicheAffaire' then
 begin //mcd 26/09/03 pour même comportement que dasn AffDispatcHtt et Ok duplication
 If (Action = TaCreat) then Lequel:='';  // on passe le code article d'origine, il faut l'effacer
 If TypeArticle ='POU' then AGLLanceFiche('AFF','AFARTPOURCENT','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle+';'+Range)
   else AGLLanceFiche('AFF','AFARTICLE','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle+';'+Range);
 end
 else if ArticleSpecif='FicheModeArt' then AGLLanceFiche('MBO','ARTICLE','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle+';'+TT)
else if ArticleSpecif='FicheModePre' then DispatchArtMode (11,TypeArticle,Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle)
else if ArticleSpecif='FicheModeFi'  then AGLLanceFiche('MBO','MBOARTFINANCIER',TypeArticle,Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle)
else AGLLanceFiche('GC','GCARTICLE','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle) ;
end;

Procedure TraiteMenusGRCModule ( NumModule : integer ) ;
begin
// Pl le 14/08/07 : FQ10622 cette restriction n'existe plus...
//  if (GetParamSocSecur ('SO_AFCLIENT', 0) <>8 ) then  //cInClientKPMG =8 pour ne pas inclure menu spécif
//    if numModule = 92 then FMenuG.RemoveGroup(92650,true); //paramètre projet

{$IF Defined(CRM) or Defined(GRCLIGHT)}
   if NumModule = 65 then { Gestion des paramètres}
{$ELSE}
   if NumModule = 92 then { Gestion des paramètres}
{$IFEND}
  begin
    if V_PGI.NumVersionSoc > 700 then
      if not GetParamsocSecur('SO_RTGESTIONGED', false) then
         if NumModule = 92 then FMenuG.RemoveGroup(92680,true);
  {$IFNDEF CRM}
    FMenuG.RenameGroup(92860,TraduireMemoire('Relation clients'),True);
  {$ENDIF !CRM}
  end;

{$IFDEF GRCLIGHT}
  if NumModule = 92 then
  begin
    if not (ctxScot in V_PGI.PGIContexte) then
    begin     //suppression des menus segmentation conseil et apporteurs pour la GC/GA 
      FMenuG.RemoveGroup(-297340, True);
      Fmenug.removeItem(297103); //pas de menu segmentation dans menu GRC
      Fmenug.removeItem(92343); //pas de menu apportuer
      Fmenug.removeItem(92258); //pas de menu suivi apporteur
      Fmenug.removeItem(92866); //pas de menu motif retrib
      Fmenug.removeItem(92380); //pas de menu suivi partenaire
      Fmenug.removeItem(92390); //pas de menu suivi des frais
    end;
//GP_20080826_DS_GP15400
    if not (ctxAffaire in V_PGI.PGIContexte) then
    begin
      FMenuG.RemoveGroup(92900,true); FMenuG.RemoveGroup(92500,true);
      FMenuG.RemoveGroup(92600,true); FMenuG.RemoveGroup(92680,true);
      FMenuG.RemoveGroup(-92913,true); FMenuG.RemoveGroup(-92609,true);
      FMenuG.RemoveGroup(92650,true) ;
    end;
  end;
{$ENDIF GRCLIGHT}

  // BDU - 25/10/07 - FQ : 13993. Ces options ne sont plus présentes en GA mais toujours en GI
  if (NumModule = 92) and (ctxAffaire in V_PGI.PGIContexte) and (not (ctxScot in V_PGI.PGIContexte)) then
  begin
    FMenuG.RemoveGroup(92650, True);
//  MODIF BRL 5/08/2010 :    FMenuG.RemoveGroup(-92913, True);
    FMenuG.RemoveGroup(-92609, True);
    FMenuG.RemoveItem(92944);
    FMenuG.RemoveItem(-92810);
    FMenuG.RemoveItem(-92830);
  end;
  if (numModule = 92) or (Nummodule = 311) then
  begin
{$IFDEF CRM}
       FMenuG.RemoveGroup(92900,True);
       FMenuG.RemoveGroup(92500,True);
       FMenuG.RemoveGroup(92600,True);
       FMenuG.RemoveGroup(92680,True);
{$ENDIF}
    if not VH_GC.GCIfDefCEGID then
    begin
      FMenuG.RemoveGroup(92800, True);
    end;
  end;
end;

end.



