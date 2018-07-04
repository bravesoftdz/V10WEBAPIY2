unit UTilSelection;

interface
                                    
uses  HEnt1, ExtCtrls, EntRT,
{$IFDEF EAGLCLIENT}
      eMul,eQRS1,MaineAGL,
	Efiche,
{$ELSE}
      QRS1,Mul,FE_Main,
  Fiche,
{$ENDIF}
{$ifdef GIGI}
        entgc,Graphics ,Vierge,
{$endif}

      Hctrls,StdCtrls, Forms,Controls,
      Classes,SysUtils,comctrls,
      UTOB,
      Stat,ParamSoc, Cube, TiersUtil,SaisieList;

const
     // Valeurs MUL :
     //  - indices de 0 à 2 : 3 lignes 3 colonnes ( entre 7 et 9 champs, sans date ou Numerique )
     //  - indices de 3 à 4 : 3 lignes 2 colonnes ( moins de 7 champs, sans date ou Numerique )
     //  - indices de 5 à 6 : 4 lignes 2 colonnes ( avec champ(s) date ou Numerique )
     //  - indices 7 : 4 lignes 1 colonne ( avec champ(s) date ou Numerique )
{$IFDEF EAGLCLIENT}
     MulChampLeft : array[0..7] of Integer = (106,315,530,160,450,99,413,200);
     MulLabelLeft : array[0..7] of Integer = (4,220,428,12,310,4,318,10);
{$ELSE}
     MulChampLeft : array[0..7] of Integer = (101,305,515,160,450,99,413,200);
     MulLabelLeft : array[0..7] of Integer = (4,215,418,12,310,4,318,10);
{$ENDIF}
     { mng 09/03/07 passage à 12 possibles
     MulNbPanelMax : integer = 9;}
     MulNbPanelMax : integer = 12;
     MulHauteurPanel : integer = 21;
     // indice 0 : 3 colonnes
     // indice 1 : 2 colonnes sans date ou Numerique ou 1 colonne avec date ou numerique
     // indice 2 : 2 colonnes avec date ou Numerique
     MulLargeurPanel  : array[0..2] of Integer = (110,140,98);
     MulLargeurLabel  : array[0..2] of Integer = (85,150,94);

     // Valeurs CUBE
     CubChampLeft : array[0..4] of Integer = (180,180,180,180,180);
     CubLabelLeft : array[0..4] of Integer = (15,15,15,15,15);
     CubNbPanelMax : integer = 15;
     CubHauteurPanel : integer = 30;
     CubLargeurPanel  : array[0..1] of Integer = (120,120);
     CubLargeurLabel  : array[0..1] of Integer = (163,163);


// Idem Cube ou Mul
     //  - indices de 0 à 2 : 3 lignes
     //  - indices de 3 à 6 : 4 lignes
     LabelTop  : array[0..6] of Integer = (9,38,68,6,28,50,72);
     ChampTop  : array[0..6] of Integer = (5,34,64,3,25,47,69);
     MulNbCol : integer = 3;
     MulNbLig : integer = 3;

// 2 colonnes / 4 lignes : si champ(s) date ou Numerique
     DNbCol : integer = 2;
     DNbLig : integer = 4;


     ChampMul : string = 'RPR_RPRLIBMUL';
     PosMul : Integer = 14;

     CodeFichierProspect : String = '0';
     CodeFichierFournisseur : String = '3';
     TabletteProspect : String = 'RTLIBCHAMPSLIBRES';
     TabletteAutres : String = 'RTLIBCHAMPS';
     NomTabletteCombo : String = 'RTRPRLIBTABLE';
     NomTabletteComboMul : String = 'RTRPRLIBTABMUL';
     ParamSocFichier : String = 'SO_RTGESTINFOS00' ;

type
     THEditGRC = class(Thedit)
         procedure TexteClick ( Sender: TObject);
     end;
var
    NbLig,NbCol : integer;
    TabletteLibelle, TabletteCombo, TabletteComboMul : String;
    ChampLeft : array[0..7] of Integer ;
    LabelLeft : array[0..7] of Integer ;
    NbPanelMax : integer ;
    HauteurPanel : integer ;
    LargeurPanel  : array[0..2] of Integer;
    LargeurLabel  : array[0..2] of Integer;
    NbOngletGrc : integer; //MCD 04/07/2005 pour connaitre le nombre d'onglet en GRC
 {$ifdef GIGI} //MCD 04/07/2005
    Hauteur : integer;//mcd 20/07/2006
 {$ENDIF GIGI}

    procedure MulCreerPanel (IsMul : Boolean; IsStat : Boolean; CodeFichier : String; NbChamps : integer; F:Tform ; OngletCourant : TTabSheet; OngletActif : integer ; NoPanel : integer; FilleTobParam : TOB; IsDateOrNum,IsDp ,IsSaisCat: boolean;TypFic:string) ;  //MCD 04/07/2005 ajout isdp  + IsSaisCat le 24/04/06 + Typfic 19/06/07
    procedure MulCreerOnglet(IsMul : Boolean; F:Tform ; NoOnglet : integer; var OngletCourant : TTabSheet;Isfiche : Boolean) ;
    procedure MulCreerPagesCL( F : TForm; FicheInfos : String;Isfiche :boolean=false);
 {$ifdef GIGI} //MCD 04/07/2005
    procedure MulCreerPagesDP( F : TForm; pTypeEnreg :string='DP';IsVerrouModif:boolean=false);
 {$ENDIF GIGI}
    Function MulWhereMultiChoix (F : TForm; Prefixe,Lien : String)  : String;

implementation

procedure MulCreerPagesCL( F : TForm; FicheInfos : String;Isfiche:boolean=false);
var i,j,NoOnglet,NbPanel,MoinsMax : integer;
    FilleTobParam,FT,TobChampsProFille,TobCl : TOB;
    Code,Select : String;
    TS,OngletCourant : TTabSheet;
    NbOnglets,Dernier_Onglet,NbChamps,First : Integer;
    InfosOnglets: array[1..40] of string;
    CodeFichier,NameFic : String;
    IsMul,IsStat,IsDateOrNum,Atraiter : Boolean ;
begin
//mcd 19/08/07 modif pour ajouter Isfiche
//passer à vrai, si on vuet la création de l'ongelt sur une fiche vierge (utiliser dans la fiche tiers personalisée GI
//utilise le même mécanisme que MuLCrerPagesDP
//par contre, dans la fct MulCreerPanel, l'utilisation de ce champ est laisser en IfdefGIGI
//si doit être généraliser, il faudra supprimer le ifdef GIG.. non fait car proche diffusion
NbOngletGrc:=0; //MCD 04/07/2005
if (F is TFQrs1) or (F is TFMul) or ( F is TFStat) or (F is TFSaisieList) then
   IsMul:=True else IsMul :=False;
// Sur les Mul je ne créé pas de comboMulti sauf pour les mailing ...
if (F is TFCube) or (F is TFQrs1) or ( F is TFStat) or ( F.Name = 'RTPROS_MAILIN_FIC')
    or ( F.Name = 'RTPROS_MUL_MAILIN') or ( F.Name = 'RTPROS_MUL_SELECT' )
    or ( F.Name = 'RTPRO_MAILIN_CONT' )then
   IsStat:=True else IsStat :=False;

if IsMul then
    begin
    for i := 0 to 7 do ChampLeft[i] := MulChampLeft[i];
    for i := 0 to 7 do LabelLeft[i] := MulLabelLeft[i];
    NbPanelMax := MulNbPanelMax;
    HauteurPanel := MulHauteurPanel;
    for i := 0 to 2 do LargeurPanel[i]  := MulLargeurPanel[i];
    for i := 0 to 2 do LargeurLabel[i]  := MulLargeurLabel[i];
    end
else
    begin
    for i := 0 to 4 do ChampLeft[i] := CubChampLeft[i];
    for i := 0 to 4 do LabelLeft[i] := CubLabelLeft[i];
    NbPanelMax := CubNbPanelMax;
    HauteurPanel := CubHauteurPanel;
    for i := 0 to 1 do LargeurPanel[i]  := CubLargeurPanel[i];
    for i := 0 to 1 do LargeurLabel[i]  := CubLargeurLabel[i];
    end;

CodeFichier:=CodeFichierProspect ;

i:=Pos('=',FicheInfos);
if i <> 0 then NameFic:=Copy(FicheInfos,i+1,length(FicheInfos)-i);

VH_RT.TobChampsPro.Load;

TobCl:=VH_RT.TobParamCl.FindFirst(['CO_TYPE','CO_ABREGE'],['RPC',NameFic],TRUE) ;
if TobCl <> Nil then
   CodeFichier:=TobCl.GetValue('CO_CODE');

if CodeFichier = CodeFichierProspect then
    begin
    TabletteLibelle:=TabletteProspect;
    TabletteCombo:=NomTabletteCombo;
    TabletteComboMul:=NomTabletteComboMul;
    end
else
    begin
    if GetParamSocSecur( ParamSocFichier+CodeFichier,false ) = False then exit;
    TabletteLibelle:=TabletteAutres+'00'+CodeFichier;
    TabletteCombo:=copy(NomTabletteCombo,1,2)+'00'+Codefichier+copy(NomTabletteCombo,6,13) ;
    TabletteComboMul:=copy(NomTabletteComboMul,1,2)+'00'+Codefichier+copy(NomTabletteComboMul,6,14) ;
    end;

NbOnglets:=TPageControl(F.FindComponent('Pages')).PageCount - 1;
First:=NbOnglets+1;
NbPanel:=0;MoinsMax:=0;NbChamps:=0;IsDateOrNum:=False;
Dernier_Onglet:= 99;

VH_RT.TobChampsPro.Load;

TobChampsProFille:=VH_RT.TobChampsProMul.FindFirst(['CO_CODE'], [CodeFichier], TRUE);
if Assigned(TobChampsProFille) then
  for i := 0 to TobChampsProFille.Detail.Count-1 do
    begin
    FilleTobParam := TobChampsProFille.Detail[i];

    Code:=FilleTobParam.GetValue('RDE_CODE');
    NoOnglet:=FilleTobParam.GetValue('RDE_ONGLET');
    Select := FilleTobParam.GetValue('RDE_CRITERESEL');

    //NoOnglet:=NoOnglet+NbOnglets;
    if FilleTobParam.GetValue('RDE_PANEL') <> 99 then
       begin
       Atraiter:=true;
       if not (Isfiche) and
        (( code = 'LIG' ) or( code = 'ESP' ) or ( Select <> 'X' )) then Atraiter:=false;
       //mcd 19/08/07 pour passer par atriater if  (( code <> 'LIG' ) and ( code <> 'ESP' ) and ( Select = 'X' )) then Atraiter:=false;
       if Atraiter then
          begin
          // Création de l'Onglet s'il n'existe pas
          if NoOnglet <> Dernier_Onglet then
            begin
               NbPanel:=0;
               Dernier_Onglet:= NoOnglet;
               Inc(NbOnglets);
               MulCreerOnglet (IsMul, F, NbOnglets,OngletCourant,Isfiche);
               Inc (NbOngletGrc);    //MCD 04/07/2005
               { compter le nombre de champs de l'onglet pour réduire ou augmenter la taille des
                 champs ( plus ou moin large si 2 ou 3 colonnes de champs )
                 Inutile en Cube }
               NbChamps:=0;
               IsDateOrNum:=False;
               MoinsMax:=0;
               if IsMul then
                   begin
                   for j:=0 to TobChampsProFille.Detail.Count-1 do
                       begin
                       FT := TobChampsProFille.Detail[j];
                       if ( FT.GetValue('RDE_ONGLET')= IntToStr(FilleTobParam.GetValue('RDE_ONGLET')) )
                          and (FT.GetValue('RDE_criteresel') = 'X') and ( FT.GetValue('RDE_code') <> 'ESP' )
                          and (FT.GetValue('RDE_code') <> 'LIG' ) and ( FT.GetValue('RDE_code') <> '' ) then
                          begin
                          Inc(NbChamps);
                          if (copy(FT.GetValue('RDE_CODE'),1,1) = 'D') or
                             (copy(FT.GetValue('RDE_CODE'),1,1) = 'V') then
                             begin
                             IsDateOrNum:=True;
                             { mng 09/03/07 passage à 12 possibles
                             if IsMul then MoinsMax:=1;}
                             if IsMul then MoinsMax:=4;
                             end;
                          end;
                       end;
                   end;
               FilleTobParam := TobChampsProFille.Detail[i];
            end;
          Inc (NbPanel);
          if NbPanel <= (NbPanelMax-MoinsMax) then
                MulCreerPanel (IsMul, IsStat, CodeFichier, NbChamps, F, OngletCourant,NbOnglets,NbPanel, FilleTobParam,IsDateOrNum,False,Isfiche,'RDE'); //MCD 04/07/2005 ajout false +IsFiche et RDE le 19/06/08
          end;
       end
    else
       InfosOnglets[NbOnglets+1]:=FilleTobParam.GetValue('RDE_NOMONGLET');
    end;
for i:=1 to 3 do
   THLabel(F.FindComponent( 'TRAC_TABLELIBRE'+IntToStr(i))).Caption:=RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE);
   //SetControlCaption('TRAC_TABLELIBRE'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE)) ;

if First <> TPageControl(F.FindComponent('Pages')).PageCount then
   begin
   {if CodeFichier <> CodeFichierProspect then
      begin
      TE:= THEdit(F.FindComponent('RDA_DESC'));
      if TE <> Nil then
         TE.Text:=CodeFichier ;
      //SetControlText ('RDA_DESC',CodeFichier);
      end;}
   for i:=First to TPageControl(F.FindComponent('Pages')).PageCount - 1 do
   TTabSheet(F.FindComponent('ONGLETLIBRE'+IntToStr(i))).Caption:=InfosOnglets[i];
   end;
if (F is TFQrs1) then
    begin // QRS1
    // positionner l'Onglet "Mise en Page" en dernier
    TS:= TTabSheet(F.FindComponent('Option'));
    if TS <> Nil then
      TS.PageIndex:=TPageControl(F.FindComponent('Pages')).PageCount - 1;
    TS:= TTabSheet(F.FindComponent('Standards'));

    TS.PageControl.ActivePage:=Nil;
    TS.PageControl.SetFocus;
    TS.PageControl.ActivePage:=TS;
    end
else
    if (F is TFMul) then
      TFMul(F).Pages.ActivePage:=Nil
    else
      if (F is TFStat) then
          TFStat(F).Pages.ActivePage:=Nil
      else
      if (F is TFFiche) then   //mcd 21/06/07
          TFFiche(F).Pages.ActivePage:=Nil
      else
        if not (F is TFSaisieList) then
          TFCube(F).Pages.ActivePage:=TTabSheet(F.FindComponent('PGeneral'));
end;



procedure MulCreerPanel (IsMul : Boolean; IsStat : Boolean; CodeFichier : String ; NbChamps : integer; F:Tform;OngletCourant : TTabSheet; OngletActif : integer ; NoPanel : integer; FilleTobParam : TOB; IsDateOrNum,IsDP, IsSaisCat : boolean;TypFic:String) ;    //MCD 04/07/2005 ajout isdp +IsSaisCat 24/04/06 +TypeFic 19/06/07
var L,L_  : TLabel;
    CC : TCheckBox;
    C  : THValComboBox;
    //MC : TListBox;
    MC : THMultiValComboBox;
    E,E_  : THEditGRC;
    N,N_  : THEdit;
    OkCase,OkCombo,OkEdit,OkMulti,OkNumEdit,OkSpace,OkLigne,OkDate : boolean;
    IndCol,IndLig,IndPanel,INum  : integer;
    //Q : TQuery;
    StNum,NomChamp,Code,TC : String;
{$ifdef GIGI}
P  : TPanel;
	lib : string;
{$endif}

begin
//mcd 19/06/07 : ilf audra supprimer tous les Ifdef GIGI si IsaisCAT
//si on veut pouvoir saisir les info prospect sur une fiche vierge classqiue pour tout le monde
//pas fait ce jour car pas assez de recul en test avant diffusion

if Isdp then    //MCD 04/07/2005
  begin
  TC:=FilleTobParam.GetValue('ADP_TYPECHAMP');
  If FilleTobParam.GetValue('ADP_TYPEINFO') ='CAT' then
     begin // si varchar 6 ou 7 et catalohue service, c'est en fait une combo
     if (FilleTobParam.GetValue('ADP_TYPECHAMP') = 'VARCHAR(6)') or
       (FilleTobParam.GetValue('ADP_TYPECHAMP') = 'VARCHAR(5)') then TC:='COMBO';
     end;
  If tabletteCombo ='YYCODENAF' then TC:='EDIT';//mcd 27/07/06 car particulier naf, il faut un thedit avec tablettte poru vori coede+llibelle
  NomChamp:=FilleTobParam.GetValue('ADP_NOMCHAMP');
  Code:=FilleTobParam.GetValue('ADP_CODE');
  end
else begin
TC:=FilleTobParam.GetValue('RDE_TYPECHAMP');
NomChamp:=FilleTobParam.GetValue('RDE_NOMCHAMP');
if CodeFichier <> CodeFichierProspect then
   //NomChamp:='RDA'+copy(NomChamp,4,length(NomChamp));
   NomChamp:='RD'+CodeFichier+'_RD'+CodeFichier+copy(NomChamp,8,length(NomChamp));
Code:=FilleTobParam.GetValue('RDE_CODE');
  end;  //MCD 04/07/2005

// calcul suivant le no panel, des no ligne et colonne
IndCol:=0; // pour xxxxLeft
IndLig:=0; // pour xxxxTop
IndPanel:=0; // xxxLargeur
NbLig := MulNbLig;
NbCol := MulNbCol;

if NbChamps > 9 then
  begin
  NbLig :=MulNbLig+1;
  NbCol :=MulNbCol+1;
  end;

if IsDateOrNum then
    begin
    IndCol := ( (NoPanel-1) div DNbLig )+5;
    IndLig := (NoPanel - ( ( (NoPanel-1) div DNbLig ) * DNbLig ) - 1)+3;
    IndPanel:=2;
    if (NbChamps < 5) then
        begin
        IndPanel:=1;
        IndCol:=IndCol+2;
        end;
    end
else
    begin
    if IsMul then
        begin
        IndCol := ( (NoPanel-1) div NbCol );
        IndLig := NoPanel - ( ( (NoPanel-1) div NbLig ) * NbLig ) - 1;
        end;
    if IsMul and (NbChamps < 7) then
       begin
       IndCol:=IndCol+3;
       IndPanel:=1;
       end;
    if IsMul and (NbChamps > 9) then
       IndLig:=IndLig+3;
   end;

OkCase:=(TC='BOOLEAN');
OkCombo:=(TC='COMBO');
OkMulti:=(TC='MULTI');
OkNumEdit:=(TC='DOUBLE');
OkLigne:=(TC='LIGNE');
OkSpace:=(TC='ESPACE');
OkDate:= (TC='DATE');

OkEdit:=((TC<>'BOOLEAN')And(TC<>'COMBO')And(TC<>'MULTI')
     And(TC<>'DOUBLE')And(TC<>'ESPACE')And(TC<>'LIGNE')And(TC<>'DATE'));

if ((Not OkCase)And(Not OkCombo)And(Not OkEdit)And
    (Not OkMulti)And(Not OkNumEdit)And(Not OkLigne)And(Not OkSpace)And(Not OkDate)) then Exit;

C:=Nil;
MC:=Nil;

{$ifdef GIGI} //mcd 24/04/2006
If OkLIgne and IsSaiscat then
   begin
   P:=TPanel.Create(F);
   P.Parent:=OngletCOurant;
   P.ParentFont:=False; P.Font.Color:=clWindowText;
   P.ParentColor:=true;
   P.Width:=TFVierge(F).width;
   P.BevelOuter:=bvRaised;
   P.Height:=6;
   p.top:=hauteur;
   P.BevelInner:=bvNone ;
   //mcd 19/06/07P.Name:='L'+IntToStr(OngletActif)+IntToStr(FilleTobParam.GetValue('ADP_PANEL'));
   P.Name:='L'+IntToStr(OngletActif)+IntToStr(FilleTobParam.GetValue(TypFic+'_PANEL'));
   P.Caption:='';
   hauteur := hauteur +6;
 end
else if (  OkSpace ) then
  begin
     L:=THLabel.Create(F);
     L.Parent:=OngletCourant;
      //mcd 19/06/07 L.Name:='L'+IntToStr(OngletActif)+IntToStr(FilleTobParam.GetValue('ADP_PANEL'));
     L.Name:='L'+IntToStr(OngletActif)+IntToStr(FilleTobParam.GetValue(typFic+'_PANEL'));
     L.Width:=320;
     L.Alignment:=TaCenter;
     L.AutoSize:=False;
     L.Height:=30;
     L.Top:=8+(NoPanel*HauteurPanel);
     l.top :=12+hauteur;
     hauteur :=Hauteur +l.height+4;
        //mcd 19/06/07 Lib := FilleTobParam.GetValue('ADP_LIBELLE');
     Lib := FilleTobParam.GetValue(TypFic+'_LIBELLE');
     L.caption := ReadTokenSt(Lib);
     l.Font.Color:=clActiveCaption;
end
else begin
{$endif}
  if OkCase then
     begin
     CC:=TCheckBox.Create(F); CC.Parent:=OngletCourant; CC.Name:=NomChamp;
     if IsMul then CC.Height:=HauteurPanel;
     //CC.Left:=ChampLeft[IndCol];
     CC.Left:=LabelLeft[IndCol];
     if IsMul then
         CC.Top:=ChampTop[IndLig]
     else
         CC.Top:=5+(NoPanel*HauteurPanel);
{$ifdef GIGI}
     If IsSaisCat then
       begin //mcd 20/07/06 pour avoir présentation correct en saisie catalogue (pas comem sur mul)
       cc.top :=5+hauteur;
       hauteur :=Hauteur +HauteurPanel;
       end;
{$endif}
     CC.Width:=(LargeurPanel[IndPanel]*2)-20; // -20 approximatif sinon empiète sur le champ colone suivante

     CC.State:=cbGrayed;
     CC.AllowGrayed:=True;
{$ifdef GIGI}
     If ISSaisCat then
      begin
			CC.State:=cbUnChecked;
      CC.AllowGrayed:=False;
			end;
{$endif}
     //CC.Alignment:= taLeftJustify;
     If IsDp then CC.Caption:=TabletteLibelle  //MCD 04/07/2005
      else CC.Caption:=RechDom(TabletteLibelle,Code,FALSE);
     CC.Anchors:=[akTop,akLeft,akRight];
     CC.BiDiMode:=bdLeftToRight;
     end
  else
     begin
     L:=THLabel.Create(F); L.Parent:=OngletCourant; L.Name:='L'+NomChamp;
     if IsMul then L.Height:=HauteurPanel;
     L.Left:=LabelLeft[IndCol];
     if IsMul then
         L.Top:=LabelTop[IndLig]
     else
         L.Top:=8+(NoPanel*HauteurPanel);
{$ifdef GIGI}
     If IsSaisCat then
       begin //mcd 20/07/06 pour avoir présentation correct en saisie catalogue (pas comem sur mul)
       l.top :=8+hauteur;
       end;
{$endif}

     L.Width:=LargeurLabel[IndPanel];

     L.AutoSize:=false;
     if ISdp then   L.Caption:=TabletteLibelle //MCD 04/07/2005
       else L.Caption:=RechDom(TabletteLibelle,Code,FALSE);
     L.Anchors:=[akTop,akLeft];
     L.Alignment:=TaLeftJustify;
     L.Autosize:=True;
     L.Align:=AlNone;
     L.BiDiMode:=bdLeftToRight;
     if OkCombo or OkMulti then
        begin
        if (not IsStat) or (OkMulti) then
            begin
            C:=THValComboBox.Create(F); C.Parent:=OngletCourant; C.Name:=NomChamp;
            C.Vide:=True;
{$ifdef GIGI}
            If IsSaisCat then C.VideString:=TraduireMemoire('<<Aucun>>');
{$endif}
            if IsMul then C.Height:=HauteurPanel;
            C.Left:=ChampLeft[IndCol];
            if IsMul then
               C.Top:=ChampTop[IndLig]
            else
               C.Top:=5+(NoPanel*HauteurPanel);
{$ifdef GIGI}
           If IsSaisCat then
             begin //mcd 20/07/06 pour avoir présentation correct en saisie catalogue (pas comem sur mul)
             c.top :=5+hauteur;
             hauteur :=Hauteur +HauteurPanel;
            //mcd 19/06/07 If GetParamsocSecur ('So_ChampObliAst',false) and (FilleTobParam.GetValue('ADP_OBLIGATOIRE')='X') then
             If GetParamsocSecur ('So_ChampObliAst',false) and (FilleTobParam.GetValue(TypFic+'_OBLIGATOIRE')='X') then
                begin  //mcd 03/01/2007 13528 GC
                C.obligatory:=true;
                end;
             end;
{$endif}
            C.Width:=LargeurPanel[IndPanel];
            C.Text:='' ;
            C.Style:=csDropDownList ;
            C.ItemIndex:=0; L.FocusControl:=C;
            C.Anchors:=[akTop,akLeft];
            end
         else
            begin
            MC:=THMultiValComboBox.Create(F);
            MC.Parent:=OngletCourant;
            MC.Name:=NomChamp;
            MC.Anchors:=[akTop,akLeft];
            if not IsMul then
                MC.Top:=5+(NoPanel*HauteurPanel)
            else
                begin
                MC.Top:=ChampTop[IndLig];
                MC.Height:=HauteurPanel;
                end;
{$ifdef GIGI}
             If IsSaisCat then
               begin //mcd 20/07/06 pour avoir présentation correct en saisie catalogue (pas comem sur mul)
               mc.top :=5+hauteur;
               hauteur :=Hauteur +mc.height;
(*        If GetParamsocSecur ('So_ChampObliAst',false) and (FilleTobParam.GetValue('ADP_OBLIGATOIRE')='X') then
          begin  //mcd 03/01/2007 13528 GC   a remettre quand OK pour  TListeBox(agl 12757)
          mc.obligatory:=true;
          end; *)
               end;
{$endif}
            MC.Left:=ChampLeft[IndCol];
            MC.Width:=LargeurPanel[IndPanel];
            L.FocusControl:=MC;
            MC.Text:='' ;
            MC.Autosize:=True;
            MC.BiDiMode:=bdLeftToRight;
            end;
        if OkCombo then
            begin
            INum:=Ord(Code[3]);
            if INum < 65 then
               INum:=StrToInt(copy(Code,3,1))
            else
               INum:=Ord(Code[3])-55;
            if IsStat then
                begin  //MCD 04/07/2005 ajout isdp
                if Isdp then MC.DataType:=TabletteCombo
                 else MC.DataType:=TabletteCombo+IntToStr(INum);
                end
            else begin //MCD 04/07/2005
                if IsDp then C.DataType:=TabletteCombo
                 else C.DataType:=TabletteCombo+IntToStr(INum);
                end;
            end
        else
            begin
            if (IsStat) and (not OkMulti) then
                begin
                if Isdp then  MC.DataType:=TabletteComboMul    //MCD 04/07/2005
                 else MC.DataType:=TabletteComboMul+copy(Code,3,1);
                //MC.Operateur:=Contient;
                end
            else
                begin
                if Isdp then C.DataType:=TabletteComboMul //MCD 04/07/2005
                 else C.DataType:=TabletteComboMul+copy(Code,3,1);
                C.Operateur:=Contient;
                end;
            end;
        if (not IsStat) or (OkMulti) then
            begin
            C.Enabled:=True;
            C.BiDiMode:=bdLeftToRight;
            end;
        end;

     if OkEdit or OkDate then
        begin
        E:=THEditGRC.Create(F); E.Parent:=OngletCourant; E.Name:=NomChamp;
        if TC='DATE' then begin E.EditMask:=DateMask; E.OpeType := otDate;E.Defaultdate:=od1900 ; end
        else E.OpeType := otString;
        if IsMul then E.Height:=HauteurPanel;

        E.Left:=ChampLeft[IndCol];
        if IsMul then
           E.Top:=ChampTop[IndLig]
        else
           E.Top:=5+(NoPanel*HauteurPanel);
{$ifdef GIGI}
       If IsSaisCat then
         begin //mcd 20/07/06 pour avoir présentation correct en saisie catalogue (pas comem sur mul)
         E.top :=5+hauteur;
         hauteur :=Hauteur +HauteurPanel;
      //mcd 19/06/07    If GetParamsocSecur ('So_ChampObliAst',false) and (FilleTobParam.GetValue('ADP_OBLIGATOIRE')='X') then
         If GetParamsocSecur ('So_ChampObliAst',false) and (FilleTobParam.GetValue(TypFic+'_OBLIGATOIRE')='X') then
           begin  //mcd 03/01/2007 13528 GC
           E.obligatory:=true;
           end;
         end;
       if TabletteCombo ='YYCODENAF' then
           begin //mcd 27/07/06 cas particulier naf, ou il faut code+ liebllé donc passer par edit avec tablette
           E.ElipsisButton:=True;
           E.Datatype:=TabletteCOmbo;
           end;
{$endif}
        E.Width:=LargeurPanel[IndPanel];
        E.Text:=''; L.FocusControl:=E;
        E.Anchors:=[akTop,akLeft];
        if TC='DATE' then E.text:=DateToStr(iDate1900) ;
        E.Autosize:=True;
        E.BiDiMode:=bdLeftToRight;
        E.Operateur:=Egal;
        if OkEdit then
           E.Operateur:=Commence;
        if OkDate then
           E.Operateur:=Superieur;
        //MCD 04/07/2005 ajout isdp
        if (not Isdp) and(OkEdit) and (FilleTobParam.GetValue('RDE_TYPETEXTE') <> '' ) then
           begin
           E.ElipsisButton:=True;
           E.OnElipsisClick:=E.TexteClick;
           E.Enabled:=true;
           if FilleTobParam.GetValue('RDE_TYPETEXTE') = 'R'  then E.Tag:=1;
           if FilleTobParam.GetValue('RDE_TYPETEXTE') = 'C'  then E.Tag:=2;
           if FilleTobParam.GetValue('RDE_TYPETEXTE') = 'O'  then E.Tag:=6;
           if FilleTobParam.GetValue('RDE_TYPETEXTE') = 'P'  then E.Tag:=7;

           if CodeFichier = CodeFichierFournisseur then E.Tag:=E.Tag+3;
           E.Hint:=FilleTobParam.GetValue('RDE_FILTRE');
           end;
        if OkDate then
            begin
{$ifdef GIGI}
				if not IsSaisCat then    //mcd 19/06/07 .. pas de date dans catalogue, donc n'était pas gerer
					begin
{$endif}
            L_:=THLabel.Create(F); L_.Parent:=OngletCourant; L_.Name:='L'+NomChamp+'_';
            L_.Height:=L.Height;
            if IsMul then
               L_.Left:=E.left+E.Width+4
            else
               L_.Left:=E.left+E.Width+20;
            L_.Top:=L.Top;
              //mcd pas de date dans le catalogue, donc pas traiter pour champ hauteur
            L_.Width:=10;
            L_.AutoSize:=false;
            L_.Caption:='à';
            L_.Anchors:=[akTop,akLeft];
            L_.Alignment:=TaLeftJustify;
            L_.Align:=AlNone;
            L_.BiDiMode:=bdLeftToRight;
            E_:=THEditGRC.Create(F); E_.Parent:=OngletCourant; E_.Name:=NomChamp+'_';
            E_.EditMask:=DateMask; E_.OpeType := otDate;E_.Defaultdate:=od1900 ;
            if IsMul then
               E_.Left:=E.left+E.Width+16
            else
               E_.Left:=E.left+E.Width+45;
            E_.Top:=E.Top;
            E_.Width:=E.Width;

            E_.Anchors:=[akTop,akLeft];
            E_.text:=DateToStr(iDate1900) ;
            E_.Autosize:=True;
            E_.BiDiMode:=bdLeftToRight;
            E_.Operateur:=Inferieur;
{$ifdef GIGI}
				end;
{$endif}

            end;
        end;
     if OkNumEdit then
        begin
        N:=THEdit.Create(F); N.Parent:=OngletCourant; N.Name:=NomChamp;
        if IsMul then N.Height:=HauteurPanel;
        N.OpeType := otReel;
        N.Left:=ChampLeft[IndCol];
        if IsMul then
           N.Top:=ChampTop[IndLig]
        else
           N.Top:=5+(NoPanel*HauteurPanel);
{$ifdef GIGI}
         If IsSaisCat then
           begin //mcd 20/07/06 pour avoir présentation correct en saisie catalogue (pas comem sur mul)
           N.top :=5+hauteur;
           hauteur :=Hauteur +HauteurPanel;
(*        If GetParamsocSecur ('So_ChampObliAst',false) and (FilleTobParam.GetValue(TypFic+'_OBLIGATOIRE')='X') then
          begin  //mcd 03/01/2007 13528 GC     a remettre quand OK pour ThNUmEdit (agl 12757)
          N.obligatory:=true;
          end; *)
           end;
{$endif}
        N.Width:=LargeurPanel[IndPanel];
        N.Text:=''; L.FocusControl:=N;
        N.Anchors:=[akTop,akLeft];
        N.Autosize:=True;
        N.BiDiMode:=bdLeftToRight;
        N.Operateur:=Superieur;
{$ifdef GIGI}
				if not IsSaisCat then
					begin
{$endif}
        // zone 2 fourchette
        L_:=THLabel.Create(F); L_.Parent:=OngletCourant; L_.Name:='L'+NomChamp+'_';
        L_.Height:=L.Height;
        if IsMul then
           L_.Left:=N.left+N.Width+4
        else
           L_.Left:=N.left+N.Width+20;
        L_.Top:=L.Top;
        L_.Width:=10;
        L_.AutoSize:=false;
        L_.Caption:='à';
        L_.Anchors:=[akTop,akLeft];
        L_.Alignment:=TaLeftJustify;
        L_.Align:=AlNone;
        L_.BiDiMode:=bdLeftToRight;

        N_:=THEdit.Create(F); N_.Parent:=OngletCourant; N_.Name:=NomChamp+'_';
        N_.Height:=N.Height;
        N_.OpeType := otReel;
        if IsMul then
           N_.Left:=N.left+N.Width+16
        else
           N_.Left:=N.left+N.Width+45;
        N_.Top:=N.Top;
        N_.Width:=N.Width;
        N_.Text:=''; L.FocusControl:=N_;
        N_.Anchors:=[akTop,akLeft];
        N_.Autosize:=True;
        N_.BiDiMode:=bdLeftToRight;
        N_.Operateur:=Inferieur;
{$ifdef GIGI}
				end;
{$endif}
        end;
     end;
{$ifdef GIGI}
end;
{$endif}

end;


procedure MulCreerOnglet (IsMul : Boolean; F:Tform ;NoOnglet : integer;var OngletCourant : TTabSheet;Isfiche:Boolean) ;
var TS,TSDim : TTabSheet ;
    i : Integer;
    B : TBevel;
begin
{$ifdef GIGI}
if IsFiche then Hauteur :=0; //mcd 19/06/07
{$endif}
// Création d'un Onglet
    TS:=TTabSheet.create(F) ;
    OngletCourant:=TS;
    With TS do
        Begin
        PageControl:=TPageControl(F.FindComponent('Pages')) ;
        i :=  NoOnglet;
        Name:='ONGLETLIBRE'+IntToStr(i);
        Parent:=TPageControl(F.FindComponent('Pages'));
        ShowHint:=FALSE ;
        //Caption:=RechDom('RTLIBONGLET','ON'+IntToStr(i-1-NbOnglets),FALSE);
        Visible:=TRUE ;
        End ;
    if IsMul then
        begin
        B:=TBevel.create(F) ;
        With B do
            Begin
            i :=  NoOnglet;
            Name:='BevelOl'+IntToStr(i);
            Parent:=TS;
            Visible:=TRUE ;
            Shape:=bsBox;
            Align:=alClient;
            Style:=bsLowered;
            Anchors:=[akLeft,akTop,akRight,akBottom];
            end;
        end
    else
        begin
        TSDim := TTabSheet(F.FindComponent('PDimension'));
        if TSDim <> Nil then
           TS.PageIndex:=TSDim.PageIndex;
        end ;
end;

procedure THEditGRC.TexteClick ( Sender: TObject);
var Ctrl : TControl;
    Retour,StArg : String;
    TE: THEdit ;
    F : TForm ;
begin
Ctrl:=TControl(Sender);
StArg:='';
if (TEdit(Ctrl).Tag = 0) or (TEdit(Ctrl).Tag = 3) then
   begin
   if TEdit(Ctrl).Text <> '' then
      StArg:= 'T_TIERS='+ TEdit(Ctrl).Text;
   if TEdit(Ctrl).Tag = 0 then
      Retour:=AGLLanceFiche ('RT', 'RTTIERS_TL',StArg , '', 'PROSPECT_MUL;FILTRE='+Ctrl.Hint)
   else
      Retour:=AGLLanceFiche ('RT', 'RFTIERS_TL',StArg , '', 'PROSPECT_MUL;FILTRE='+Ctrl.Hint);
   if Retour <> '' then
       TEdit(Ctrl).Text:=Retour;
   end
else
   if (TEdit(Ctrl).Tag = 1) or (TEdit(Ctrl).Tag = 4) then
       begin
       if TEdit(Ctrl).Text <> '' then
          StArg:= 'ARS_RESSOURCE='+ TEdit(Ctrl).Text;
       Retour:=AGLLanceFiche ('RT', 'RTRESSOURCE_TL',StArg , '', 'PROSPECT_MUL;FILTRE='+Ctrl.Hint);
       if Retour <> '' then
           TEdit(Ctrl).Text:=Retour;
       end
{$IFNDEF GRCLIGHT}       
   else
     if (TEdit(Ctrl).Tag = 6) or (TEdit(Ctrl).Tag = 9) then
         begin
         if TEdit(Ctrl).Text <> '' then
            StArg:= 'ROP_OPERATION='+ TEdit(Ctrl).Text;
         if (TEdit(Ctrl).Tag = 6) then
           Retour:=AGLLanceFiche ('RT', 'RTOPERATIONS_TL',StArg , '', 'PROSPECT_MUL;FILTRE='+Ctrl.Hint)
         else
           Retour:=AGLLanceFiche ('RT', 'RFOPERATIONS_TL',StArg , '', 'PROSPECT_MUL;FILTRE='+Ctrl.Hint);  
         if Retour <> '' then
             TEdit(Ctrl).Text:=Retour;
         end
     else
       if (TEdit(Ctrl).Tag = 7) then
           begin
           if TEdit(Ctrl).Text <> '' then
              StArg:= 'RPJ_PROJET='+ TEdit(Ctrl).Text;
           Retour:=AGLLanceFiche ('RT', 'RTPROJETS_TL',StArg , '', 'PROSPECT_MUL;FILTRE='+Ctrl.Hint);
           if Retour <> '' then
               TEdit(Ctrl).Text:=Retour;
           end
{$ENDIF GRCLIGHT}
       else
       begin
       Retour:=AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=CLI','','PROSPECT_MUL');
       if Retour <> '' then
          TEdit(Ctrl).Text:=copy(Retour,1,(pos(';',Retour)-1));
       F:=TForm(TEdit(Ctrl).owner);

         TE:= THEdit(F.FindComponent('RAC_TIERS'));
       if TE <> Nil then
          TE.Text:=TiersAuxiliaire(copy(Retour,pos(';',Retour)+1,length(Retour)),true) ;
       TE:= THEdit(F.FindComponent('T_TIERS'));
       if TE <> Nil then
          TE.Text:=TiersAuxiliaire(copy(Retour,pos(';',Retour)+1,length(Retour)),true) ;

       end
;
end;
{$ifdef GIGI} //MCD 04/07/2005
procedure MulCreerPagesDP( F : TForm; pTypeEnreg :string='DP';IsVerrouModif:boolean=false);
var i,j,NoOnglet,NbPanel,MoinsMax : integer;
    FilleTobParam,FT : TOB;
    Code,Select,St : String;
    TS,OngletCourant : TTabSheet;
    NbOnglets,Dernier_Onglet,NbChamps,First,Boucle,Boucle2 : Integer;
    InfosOnglets: array[1..20] of string;
    CodeFichier: String;
    IsMul,IsStat,IsDateOrNum ,IsSaisCat, Atraiter: Boolean ;
begin
Hauteur :=0;
if (f is TfVierge) and (F.Name = 'AFCATALOGUE' )then isSaisCat:=true
   else IsSaisCat :=False;

if (F is TFQrs1) or (F is TFMul) or ( F is TFStat) or (F is TFSaisieList) then
   IsMul:=True else IsMul :=False;
// Sur les Mul je ne créé pas de comboMulti sauf pour les mailing ...
if (F is TFCube) or (F is TFQrs1) or ( F is TFStat) or ( F.Name = 'RTPROS_MAILIN_FIC')
    or ( F.Name = 'RTPROS_MUL_MAILIN') or ( F.Name = 'RTPROS_MUL_SELECT' )
    or ( F.Name = 'RTPRO_MAILIN_CONT' )then
   IsStat:=True else IsStat :=False;

if IsMul then
    begin
    for i := 0 to 7 do ChampLeft[i] := MulChampLeft[i];
    for i := 0 to 7 do LabelLeft[i] := MulLabelLeft[i];
    NbPanelMax := MulNbPanelMax;
    HauteurPanel := MulHauteurPanel;
    for i := 0 to 2 do LargeurPanel[i]  := MulLargeurPanel[i];
    for i := 0 to 2 do LargeurLabel[i]  := MulLargeurLabel[i];
    end
else
    begin
    for i := 0 to 4 do ChampLeft[i] := CubChampLeft[i];
    for i := 0 to 4 do LabelLeft[i] := CubLabelLeft[i];
    NbPanelMax := CubNbPanelMax;
    HauteurPanel := CubHauteurPanel;
    for i := 0 to 1 do LargeurPanel[i]  := CubLargeurPanel[i];
    for i := 0 to 1 do LargeurLabel[i]  := CubLargeurLabel[i];
    end;
NbOnglets:=TPageControl(F.FindComponent('Pages')).PageCount - 1;
First:=NbOnglets+1;
NbPanel:=0;MoinsMax:=0;NbChamps:=0;IsDateOrNum:=False;
Dernier_Onglet:= 99;

If pTypeEnreg = 'DP' then boucle := Vh_RT.TobChampsDpMul.detail.count -1
   else boucle := VH_GC.AfTOBCatalogue.detail.count -1;
For i:=0 to boucle do
    begin
    If pTypeEnreg = 'DP' then FilleTobParam := Vh_RT.TobChampsDpMul.detail[i]
     else FilleTobParam := VH_GC.AfTOBCatalogue.detail[i];
    Code:=FilleTobParam.GetValue('ADP_CODE');
    NoOnglet:=FilleTobParam.GetValue('ADP_ONGLET') + NbOngletGrc;
    Select := FilleTobParam.GetValue('ADP_CRITERESEL');
    St := FilleTobParam.GetValue('ADP_LIBELLE');
    TabletteLibelle:=ReadTokenSt(St);
    TabletteCombo:=FilleTobParam.GetValue('ADP_TABLETTE');
    TabletteComboMul:=FilleTobParam.GetValue('ADP_TABLETTE');
    if FilleTobParam.GetValue('ADP_PANEL') <> 99 then
       begin
			 Atraiter:=true;  //mcd 24/04/06 pour cas particulier saisie catalogue
       If Not IsSaisCat and ( ( code = 'LIG' ) or( code = 'ESP' ) or ( Select <> 'X' )) then Atraiter:=False;
      		//mcd 24/04/06  if ( code <> 'LIG' ) and ( code <> 'ESP' ) and ( Select = 'X' ) then
       if Atraiter then
          begin
          // Création de l'Onglet s'il n'existe pas
          if NoOnglet <> Dernier_Onglet then
             begin
             Hauteur :=0; //ilf aut repartir à 0 pour nouvel onglet
             NbPanel:=0;
             Dernier_Onglet:= NoOnglet;
             Inc(NbOnglets);
             MulCreerOnglet (IsMul, F, NbOnglets,OngletCourant,IsSaisCat);
             if IsVerrouModif then OngletCourant.enabled:=False; //mcd 27/06/06
             { compter le nombre de champs de l'onglet pour réduire ou augmenter la taille des
               champs ( plus ou moin large si 2 ou 3 colonnes de champs )
               Inutile en Cube }
             NbChamps:=0;
             IsDateOrNum:=False;
             MoinsMax:=0;
             if IsMul then
                 begin
                 If pTypeEnreg = 'DP' then boucle2 := Vh_RT.TobChampsDpMul.detail.count -1
                  else boucle2 := VH_GC.AfTOBCatalogue.detail.count -1;
                 for j:=0 to boucle2 do
                     begin
                     If pTypeEnreg = 'DP' then FT := VH_RT.TobChampsDPMul.Detail[j]
                      else  FT := VH_GC.AfTOBCatalogue.detail[j];
                     if ( FT.GetValue('ADP_ONGLET')= IntToStr(FilleTobParam.GetValue('ADP_ONGLET')) )
                        and (FT.GetValue('ADP_criteresel') = 'X') and ( FT.GetValue('ADP_code') <> 'ESP' )
                        and (FT.GetValue('ADP_code') <> 'LIG' ) and ( FT.GetValue('ADP_code') <> '' ) then
                        begin
                        Inc(NbChamps);
                        if (FT.GetValue('ADP_TYPECHAMP') = 'DATE') or
                           (FT.GetValue('ADP_TYPECHAMP') = 'DOUBLE') then
                           begin
                           IsDateOrNum:=True;
                             //mcd 08/02/2008 14972 aligne sur GRC if IsMul then MoinsMax:=1;}
                           if IsMul then MoinsMax:=4;
                           end;
                        end;
                     end;
                 end;
             If pTypeEnreg = 'DP' then FilleTobParam := VH_RT.TobChampsDPMul.Detail[i]
              else FilleTobParam := VH_GC.AfTOBCatalogue.Detail[i]
             end;  //fin création onglet
          Inc (NbPanel);
          if NbPanel <= (NbPanelMax-MoinsMax) then MulCreerPanel (IsMul, IsStat, CodeFichier, NbChamps, F, OngletCourant,NbOnglets,NbPanel, FilleTobParam,IsDateOrNum ,true,ISSaisCat,'ADP');
          end;  //fin zone à ajouter
       end   //fin panel <>99
    else
       InfosOnglets[NbOnglets+1]:=FilleTobParam.GetValue('ADP_NOMONGLET');
    end;

if First <> TPageControl(F.FindComponent('Pages')).PageCount then
   begin
   for i:=First to TPageControl(F.FindComponent('Pages')).PageCount - 1 do
   TTabSheet(F.FindComponent('ONGLETLIBRE'+IntToStr(i))).Caption:=InfosOnglets[i];
   end;
if (F is TFQrs1) then
    begin // QRS1
    // positionner l'Onglet "Mise en Page" en dernier
    TS:= TTabSheet(F.FindComponent('Option'));
    if TS <> Nil then
      TS.PageIndex:=TPageControl(F.FindComponent('Pages')).PageCount - 1;
    TS:= TTabSheet(F.FindComponent('Standards'));

    TS.PageControl.ActivePage:=Nil;
    TS.PageControl.SetFocus;
    TS.PageControl.ActivePage:=TS;
    end
else
    if (F is TFMul) then
      TFMul(F).Pages.ActivePage:=Nil
    else
      if (F is TFStat) then
          TFStat(F).Pages.ActivePage:=Nil
      else
        if not (F is TFSaisieList) then
          if (F is TFCube) then TFCube(F).Pages.ActivePage:=TTabSheet(F.FindComponent('PGeneral'));
end;
 {$ENDIF GIGI}

 Function MulWhereMultiChoix (F : TForm; Prefixe,Lien : String)  : String;
var ChpMulti:THMultiValComboBox;
    stListeCodes,stlieur,stCode,StWhere: String;
    i : integer;
    first : boolean;
begin
 //Where (XX_TOTO like  "01;% " OR XX_TOTO like  "%;01;% ")
  StWhere:='';
  Result:='';
  first:=true;
  for i:=0 to 19 do
    begin
    ChpMulti:=THMultiValComboBox(F.FindComponent(Prefixe+ChampMul+intToStr(i)));
    if Assigned(ChpMulti) then
      if ( ChpMulti.text <> '') and (ChpMulti.text <> TraduireMemoire('<<Tous>>')) then
        begin
        stListeCodes:=ChpMulti.text;
        Repeat
          stCode:=ReadToKenSt(stListeCodes);
          if stCode <> '' then
            begin
            if first then stlieur:='' else stLieur:=' '+Lien;
            first:=false;
            StWhere:=stLieur+' ( ('+Prefixe+'_'+Prefixe+ChampMul+intToStr(i)+' LIKE "'+stCode+';%") OR ('
              +Prefixe+'_'+Prefixe+ChampMul+intToStr(i)+' LIKE "%;'+stCode+';%")) ';
            Result:=Result+stWhere;
            end;
        until stCode = '';
        end;
    end;
    if Result <> '' then
      Result := '('+Result+')';
end;

Initialization

end.




