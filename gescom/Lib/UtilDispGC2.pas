unit UtilDispGC;

interface

Uses hent1,entGC,hctrls, sysutils,math, AglInit,
{$IFDEF EAGLCLIENT}
     MenuOLX,Maineagl,
{$ELSE}
     MenuOLG,Fe_Main,
{$ENDIF}
     UtilArticle;

  procedure AjusteMenu ( NumModule : integer ) ;
  Procedure AfterProtecGC ( sAcces : String ) ;
  procedure ChargeModules ( Charge : Boolean ) ;
  Procedure DispatchTTArticle ( Action : TActionFiche ; Lequel,TT,Range : String ) ;
// Spécificités marché MODE
  procedure ChargeModules_Mode ( Charge : Boolean ) ;
  Procedure AfterProtecMode ( sAcces : String ) ;

{$IFDEF AFFAIRE}
{$ENDIF}
{$IFDEF GRC}
{$ENDIF}

Const
{$IFDEF MODE}
     GCCodeDomaine = '30330010' ;
     GCCodesSeria:array[1..6] of string =('30250010','30260010','30270010','30280010','30290010','30300010');
     GCTitresSeria:array[1..6] of string =('CEGID Back Office S5','CEGID Gestion Commerciale S5','Licence 1 à 10 établissements S5','Licence 11 à 20 établissements S5','Licence 21 à 50 établissements S5','Licence + de 50 établissements S5');
 {$ELSE}
  {$IFDEF GRC}
    {$IFDEF AFFAIRE}
     GCCodeDomaine = '05980012' ;
     GCCodesSeria:array[1..4] of string =('00100012','00013012','05003012','00099012');
     GCTitresSeria:array[1..4] of string =('Gestion commerciale','Cegid e-commerce','Gestion de la Relation Client','Cegid Gestion d''affaires');
    {$ELSE}
      {$IFDEF CCS3}
       GCCodeDomaine = '03980012' ;
       GCCodesSeria : array[1..1] of string =('00100012');
       GCTitresSeria : array[1..1] of string =('Gestion commerciale');
      {$ELSE}
       GCCodeDomaine = '05980012' ;
       GCCodesSeria:array[1..3] of string =('00100012','00013012','05003012');
       GCTitresSeria:array[1..3] of string =('Gestion commerciale','Cegid e-commerce','Gestion de la Relation Client');
      {$ENDIF}
    {$ENDIF}
  {$ELSE}
    {$IFDEF AFFAIRE}
    GCCodeDomaine = '05980012' ;
    GCCodesSeria:array[1..3] of string =('00100012','00013012','00099012');
    GCTitresSeria:array[1..3] of string =('Gestion commerciale','Cegid e-commerce','Cegid Gestion d''affaires');
    {$ELSE}
    GCCodeDomaine = '05980012' ;
    GCCodesSeria:array[1..2] of string =('00100012','00013012');
    GCTitresSeria:array[1..2] of string =('Gestion commerciale','Cegid e-commerce');
    {$ENDIF}
  {$ENDIF}
 {$ENDIF}

{ Modules :                    59 e-commerce
30 ventes                      60 administration
31 achats                      65 Paramètres
32 Articles et stocks          70 affaires
33 analyses                    92 GRC
}
Const
{$IFDEF CCS3}
      GCModules:array[0..5] of integer =(30,31,32,33,65,60);
      GCImages:array[0..5] of integer  =(35,18, 5,11,28,49);
{$ELSE}
      GCModules:array[0..8] of integer =(30,31,32,33,59,70,92,65,60);
      GCImages:array[0..8] of integer  =(35,18, 5,11,32, 8,77,28,49);
{$ENDIF}
      ModulesCommuns=[30,31,32,33,65,60] ;
      ImagesCommuns= [35,18, 5,11,28,49] ;
      ModuleGRC=[92];       ImageGRC=[77];
      ModuleECommerce=[59]; ImageECommerce=[32];
      ModuleAffaire=[70];   ImageAffaire=[8];

implementation

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


procedure ChargeModules ( Charge : Boolean ) ;
var  LesModules,LesImages : array of integer;
     i,j:integer;
     OkOk:Boolean;
begin
SetLength(LesModules,High(GCModules)+1);
SetLength(LesImages,High(GCModules)+1);
J:=0;
For i:= 0 to High(GCModules) do
    begin
    OkOK:=False;
{$IFDEF GRC}
    if V_PGI.VersionDemo then OkOK:=True
{$ELSE}
    if V_PGI.VersionDemo and (not(GCModules[i] in ModuleGRC)) then OkOK:=True
{$ENDIF}
    else If (GCModules[i] in ModulesCommuns) then OkOk:=True
{$IFNDEF CCS3}
    else If (GCModules[i] in ModuleGRC)       and (VH_GC.GRCSeria) then OkOk:=True
    else If (GCModules[i] in ModuleECommerce) and (VH_GC.ECSeria)  then OkOk:=True
    else If (GCModules[i] in ModuleAffaire)  then OkOk:=True
{$ENDIF}
    ;
    {$IFDEF EAGLCLIENT}
    if GCModules[i]=60 then Okok:=False ; 
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
var i : integer ;

       Procedure GCSerialise(CodeSeria : string; bSeria : boolean) ;
          begin
          if (CodeSeria='00100') then  VH_GC.GCSeria :=bSeria else
          if (CodeSeria='00013') then  VH_GC.ECSeria :=bSeria else
          if (CodeSeria='05003') then  VH_GC.GRCSeria:=bSeria else
          if (CodeSeria='00099') then  VH_GC.GASeria :=bSeria ;
          end;
BEGIN
VH_GC.GCSeria:=False  ; VH_GC.ECSeria:=False ;
VH_GC.GRCSeria:=False ; VH_GC.BOSeria :=False;
VH_GC.GASeria :=False ;

if ctxMode in V_PGI.PGIContexte then begin AfterProtecMode (sAcces); exit; end;

if Length(sAcces)>0 then for i:=1 to Length(sAcces) do GCSerialise(copy(GCCodesSeria[i],1,5), (sAcces[i]='X') ) ;
V_PGI.VersionDemo:=((not VH_GC.GCSeria) or V_PGI.NoProtec) ;
{$IFDEF CEGID}
V_PGI.VersionDemo:=False;
VH_GC.GCSeria:=True  ; VH_GC.GRCSeria:=True ; VH_GC.GASeria :=True ;
{$ENDIF}

{$IFDEF GRC}
if V_PGI.VersionDemo then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC]
else if VH_GC.GRCSeria then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC] ;
{$ENDIF}

{$IFDEF AFFAIRE}
if (V_PGI.VersionDemo) then VH_GC.GASeria:=True;
{$ENDIF}

ChargeModules(True) ;
END ;

//=======================================================================================
//+++++++++++++++++++ Spécificité suivant marché ++++++++++++++++++++++++++++++++++++++++
//=======================================================================================

// Procédure de chargement des modules Mode dans le menu, en fonction de la Séria
procedure ChargeModules_Mode ( Charge : Boolean ) ;
BEGIN
{$IFDEF EAGLCLIENT}
{$ELSE}
if not (ctxMode in V_PGI.PGIContexte) then exit;
TitreHalley:='CEGID MODE S5';
if Charge=True then FMenuG.SetModules([101,102,103,59,111,110,105,106],[116,94,121,94,69,1,74,78]);
if (V_PGI.VersionDemo = True) then
    begin
    VH_GC.NbEtablisSeria := 4;
    TitreHalley:='CEGID MODE S5 (DEMO)';
    end else
    begin
    if (VH_GC.BOSeria = False) then
       begin
       TitreHalley:='CEGID GESTION COMMERCIALE MODE S5';
       end;
    if (VH_GC.GCSeria = False) then
       begin
       TitreHalley:='CEGID BACK OFFICE MODE S5';
       if Charge=True then FMenuG.SetModules([101,115,103,59,111,110,105,106],[116,94,121,94,69,1,74,78]);
       end;
    end;
FMenuG.UpdateCaption('');
V_PGI.NbColModuleButtons:= 2;
V_PGI.NbRowModuleButtons:= 4;
{$ENDIF}
END;

Procedure AfterProtecMode ( sAcces : String ) ;
Var BOSeria,GCSeria : boolean;
BEGIN
if not (ctxMode in V_PGI.PGIContexte) then exit;
BOSeria:=False ;
GCSeria:=False ;
VH_GC.NbEtablisSeria := 0;
if Length(sAcces)>0 then
   begin
   BOSeria:=(sAcces[1]='X') ;  // Back Office Mode sérialisé
   GCSeria:=(sAcces[2]='X') ;  // Gestion Commerciale Mode sérialisé
   // Licence Nombre Etablissements sérialisé
   if (sAcces[3]<>'X') then VH_GC.NbEtablisSeria := 0
   else if (sAcces[4]<>'X') then VH_GC.NbEtablisSeria := 10
   else if (sAcces[5]<>'X') then VH_GC.NbEtablisSeria := 20
   else if (sAcces[6]<>'X') then VH_GC.NbEtablisSeria := 50
   else VH_GC.NbEtablisSeria := 9999;
   end;
   // cas ou on coche "version non sérialiser' dans
   // l'écran de séria. On force la non sérialisation
  if V_PGI.NoProtec then
     begin
     V_PGI.VersionDemo := True ;
     BOSeria:=False ;
     GCSeria:=False ;
     end
     else V_PGI.VersionDemo := False;

  if (BOSeria=False) and (GCSeria=False) then  V_PGI.VersionDemo := True; // on repasse en demo dans ce cas
  if (VH_GC.NbEtablisSeria=0) then V_PGI.VersionDemo := True;

  VH_GC.BOSeria := BOSeria;
  VH_GC.GCSeria := GCSeria;

  ChargeModules_Mode ( True );
END ;

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
if ArticleSpecif='FicheAffaire'
   then AGLLanceFiche('AFF','AFARTICLE','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle)
else if ArticleSpecif='FicheModeArt'
   then AGLLanceFiche('MBO','ARTICLE','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle)
else if ArticleSpecif='FicheModePre'
   then AGLLanceFiche('MBO','ARTPRESTATION',TypeArticle,Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle)
else if ArticleSpecif='FicheModeFi'
   then AGLLanceFiche('MBO','ARTFINANCIER',TypeArticle,Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle)
else AGLLanceFiche('GC','GCARTICLE','',Lequel,ActionToString(Action)+';MONOFICHE;TYPEARTICLE='+TypeArticle) ;
end;

{$IFDEF AFFAIRE}
{$ENDIF}
{$IFDEF GRC}
{$ENDIF}

end.

