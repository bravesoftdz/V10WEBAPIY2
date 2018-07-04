{***********UNITE*************************************************
Auteur  ...... : AB 
Créé le ...... : 10/06/2002
Modifié le ... : 10/06/2002
Description .. : Source Saisie du BUDGET
Mots clefs ... : TOF;BUDGET
*****************************************************************}
Unit UTOFAFBUDGET ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils,ComCtrls,Messages,Windows,
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   dbTables, db,FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox,HPanel, UTOF ,vierge,HTB97,
     UTOB,Affaireutil,AglInitGC,M3FP,Saisutil,Grids,
     DicoAF,graphics,EntGC, Ent1,CalcOleGenericAff,
     utilressource,AfUtilArticle,Utilarticle ;

Type
  TOF_AFBUDGET = Class (TOF)
  public
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnDelete                 ; override ;
    procedure OnClose                  ; override ;

  private
    TobAff,TobAffs,TOBGdBudget,TOBLigne,TOBL,TobAfBudget: TOB;
    Titre,StAffaire,LesColonnes,LesMois,TypeBudg : string ;
    NombreMois :integer;
    datedeb,datefin :Tdatetime;
    Erreur : boolean;
    GS : THGRID ;
    ColTitre : TStringList;
    Action   : TActionFiche ;
    Devise : RDEVISE;
    BNEWLINE,BDELLINE,BCUMUL: TToolbarButton97;
    PSAICUMUL : THPanel;
    CUMBUDGET : THNumEdit;
    procedure FermeTout ;
    procedure OuvreTout ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SupprimeLigne(ARow:integer) ;
    procedure InsertLigne(var ARow:integer) ;
    procedure GSElipsisClick(Sender: TObject);
    procedure BNEWLINEClick(Sender: TObject);
    procedure BDELLINEClick(Sender: TObject);
    procedure BCUMULClick(Sender: TObject);
    procedure GSRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit  (Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSComboPopUp(Sender : TObject);
    procedure GSDessineCell( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
    procedure ChargeLescolonnes;
    procedure ChargeGridBudget;
    procedure AffecteCumul ;
    function  ChargeBudget (Tobbudg:TOB):boolean ;
    procedure TraiteBudget ;
    procedure ValideBudget ;


  end ;

Const C_Objects   : integer = 0;
      C_type      : integer = 1;
      C_Article   : integer = 2;
      C_Generation: integer = 3;
      C_cumul     : integer = 4;
      C_mois      : integer = 5;

// libellés des messages de la TOF  Afbudget
TexteMsgAffaire: array[1..8] of string 	= (
  {1}        'Liste des articles'
  {2}        ,'Vous devez renseigner la date de début de facturation de l''affaire'
  {3}        ,'Vous n''avez pas renseigné la date de fin de facturation de l''affaire'
  {4}        ,'ATTENTION : Budget non enregistré.'
  {5}        ,'ATTENTION : Ce Budget en cours de traitement par un autre utilisateur n''a pas été enregistrée.'
  {6}        ,'Confirmez-vous la suppression du budget ?'
  {7}        ,'Votre période de facturation différe des dates de votre budget'
  {8}        ,'Le code article est invalide'
             );

Implementation

procedure TOF_AFBUDGET.OnArgument (StArgument : String ) ;
  Var  Critere, Champ, valeur  : String;
    x : integer;
    zaff,libcli,st,tlib,lib1,lib2,debfact,finfact: string;
begin
  Inherited ;
  // Recup des critères
  Critere:=(Trim(ReadTokenSt(stArgument)));
  Action:=taModif ; Erreur:=False;
  While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos('=',Critere);
        if x=0 then X:=pos(':',Critere); // mcd 08/04/03 : dans la fiche .. on traite les 2 cas ...
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'ZAFF' then Zaff := Valeur;
        if Champ = 'ZDEB' then debfact := Valeur;
        if Champ = 'ZFIN' then finfact := Valeur;
        if (Champ ='ACTION') and (valeur = 'CONSULTATION') then Action:=taConsult;
        END;
  Critere:=(Trim(ReadTokenSt(stArgument)));
  END;

  OuvreTout;

  AfRemplirTobAffaire(zaff,TobAff,TobAffs);

  libcli := RecupLibelleTiers(TOBAff.GetValue('AFF_TIERS'));
  st := CodeAffaireAffiche(zaff,' ');

  tlib := format('%s / %s   %s',[libcli,st,TOBAff.GetValue('AFF_LIBELLE')]);
  SetcontrolText('TLIBCLIAFF',tlib);

  SetControlText('AFF_AFFAIRE',TOBAff.GetValue('AFF_AFFAIRE'));
  SetControlText('AFF_TIERS',TOBAff.GetValue('AFF_TIERS'));
  stAffaire := TOBAff.GetValue('AFF_AFFAIRE');
  lib1:='';lib2:='';
  LibelleRessource( TOBAff.GetValue('AFF_RESPONSABLE'),lib1,lib2);
  st := format('%s %s',[lib1,lib2]);
  SetControlText('TAFF_RESPONSABLE_',st);
  if (IsValidDate(debfact)) then datedeb := StrToDate(debfact)
  else datedeb := TOBAff.GetValue('AFF_DATEDEBGENER');
  if (IsValidDate(finfact)) then datefin := StrToDate(finfact)
  else datefin := TOBAff.GetValue('AFF_DATEFINGENER');

  SetControlText('TAFF_DATEDEBGENER',datetostr(datedeb));
  SetControlText('TAFF_DATEFINGENER',datetostr(datefin));

  THValComboBox(GetControl('ABU_TYPEAFBUDGET')).ItemIndex:=1;
  Typebudg := 'PVT';
  BDELLINE:=TToolbarButton97(GetControl('BDELLINE'));
  BDELLINE.OnClick:=BDELLINECLick;
  BDELLINE.visible :=(Action<>taConsult);
  BNEWLINE:=TToolbarButton97(GetControl('BNEWLINE'));
  BNEWLINE.OnClick:=BNEWLINECLick;
  BNEWLINE.visible :=(Action<>taConsult);
  BCUMUL:=TToolbarButton97(GetControl('BCUMUL'));
  BCUMUL.OnClick:=BCUMULCLick;
  BCUMUL.visible :=(Action<>taConsult);
  PSAICUMUL:=THPANEL(GetControl('PSAICUMUL'));  
  GS:= THGRID(GetControl('GRIDBUDG'));
  GS.ColFormats[C_type]:='CB=GCTYPEARTICLE||';
  GS.ColFormats[C_Generation]:='CB=AFBUDGETECHE||';
  GS.OnRowEnter:=GSRowEnter ;
  GS.OnRowExit:=GSRowExit;
  GS.OnCellEnter:=GSCellEnter ;
  GS.OnCellExit:=GSCellExit ;
  GS.OnElipsisClick:=GSElipsisClick ;
  GS.PostDrawCell := GSDessineCell;
  AffecteGrid(GS,Action) ;
  TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
  Titre := TFVierge(Ecran).caption;
end ;

procedure TOF_AFBUDGET.OnLoad ;
var datePlus24 :TDateTime;
    i :integer;
    Chargement :boolean;
begin
  Inherited ;
  titre:=ecran.caption; Erreur:=False;  CUMBUDGET:=nil;
                                                                
  if (datedeb=iDate1900) then
  begin
    PGIBoxAf(TexteMsgAffaire[2],Titre);
    PostMessage(TFVierge(Ecran).Handle,WM_CLOSE,0,0) ; exit;
  end;

//  if (datefin>datePlus24 ) then datefin := datePlus24; gm   18/07/02     
   if (datefin=idate2099 ) then
   Begin
    PGIBoxAf(TexteMsgAffaire[3],Titre);
   	datePlus24 := FinDeMois(PlusMois (datedeb,23));
   	datefin := datePlus24;
   End;

  TOBLigne.addchampsupValeur('TYPE','',false);
  TOBLigne.addchampsupValeur('CODE','',false);
  TOBLigne.addchampsupValeur('GENE','EC1',false);
  TOBLigne.addchampsupValeur('CUMUL',0,false);
  LesColonnes := 'TYPE;CODE;GENE;CUMUL';
  ColTitre := TStringList.Create;
  ColTitre.Add(' '); ColTitre.Add('Type');  ColTitre.Add('Article');
  ColTitre.Add('Alim Ech.');  ColTitre.Add('Mt Cumul');
  NombreMois := 0; LesMois:='';

  Chargement := ChargeBudget(TobAfBudget);
  ChargeLescolonnes;
  TOBLigne.addchampsupValeur('ARTICLE','',false);

  GS.colcount:=C_mois+NombreMois;
  GS.Titres := ColTitre; ColTitre.clear;ColTitre.Free;

  GS.colwidths[C_type]:= 80; GS.colwidths[C_Article]:= 80; GS.colwidths[C_Generation]:= 70;
  for i:=C_cumul to NombreMois+C_cumul do
  begin
    GS.ColTypes[i]:='R'; GS.ColFormats[i]:='##0';
    GS.colwidths[i]:= 47;    GS.ColAligns[i]:=taRightJustify;
  end;
  GS.colwidths[C_cumul]:= 57;
  if Chargement then ChargeGridBudget
  else begin
    Tob.create ('Ligne budget',TOBGdBudget,-1);
    TOBGdBudget.detail[0].Dupliquer (TOBLigne, False, True, True) ;
  end;
  TOBGdBudget.PutGridDetail(GS,false,false,LesColonnes,False);
  TOBGdBudget.SetAllModifie (false);
  GS.rowcount := TOBGdBudget.detail.count+1;
  if (NombreMois < 9) then
  begin
    if (GS.ScrollBars=ssBoth) then GS.ScrollBars:=ssVertical else GS.ScrollBars:=ssNone;
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS);
    if (NombreMois >6) then
    GS.colwidths[C_cumul]:= GS.colwidths[C_cumul]-3;
  end;
  AffecteCumul;
  Devise.Code := TobAff.GetValue ('AFF_DEVISE');
  if Devise.code = '' then Devise.Code:= V_PGI.DevisePivot;
  GetInfosDevise(Devise);
  GS.SetFocus;
  GS.Col := C_type; GS.Row := GS.rowcount-1;
  GS.ShowCombo(GS.Col,GS.Row) ;
  GS.valcombo.OnDropDown := GSComboPopup;
end ;

procedure TOF_AFBUDGET.ChargeLescolonnes;
var encours,mois :TDateTime;
    i_ind :integer;
    champ :string;
    periodechange :boolean;
begin
  mois:= iDate1900; NombreMois :=0; periodechange:=False;
 if TobAfBudget.detail.count >0 then
 begin
    encours := datedeb;
    mois := TobAfBudget.detail[0].getvalue('ABU_DATEBUD');
    if (mois <> DebutDeMois(datedeb)) then periodechange := true;
    while (encours <mois ) do
    begin
      inc(NombreMois); GS.colcount:= C_mois+NombreMois;
      champ := formatDateTime('mm/yy',encours);
      ColTitre.Add (formatDateTime('mmmyy',encours));
      LesMois := LesMois+champ+';' ;
      LesColonnes := LesColonnes+';'+ champ;
      TOBLigne.addchampsupvaleur (champ,0,false);
      encours := DebutDeMois(PlusMois (encours,1));
    end;
  end;

  for i_ind := 0 to TobAfBudget.detail.count-1 do
  begin
    encours := TobAfBudget.detail[i_ind].getvalue('ABU_DATEBUD');
    if (encours < mois)  then break;
    inc(NombreMois);
    mois := encours; GS.colcount:= C_mois+NombreMois;
    champ := formatDateTime('mm/yy',encours);
    ColTitre.Add (formatDateTime('mmmyy',encours));
    LesMois := LesMois+champ+';' ;
    LesColonnes := LesColonnes+';'+ champ;
    TOBLigne.addchampsupvaleur (champ,0,false);
  end;
  if (NombreMois > 0) then
  begin
    if (mois <> DebutDeMois(datefin)) then periodechange := true;
    encours := DebutDeMois(PlusMois (mois,1));
  end else
    encours := DebutDeMois(datedeb);

  while (encours <= datefin) do
  begin
    inc(NombreMois); GS.colcount:= C_mois+NombreMois;
    champ := formatDateTime('mm/yy',encours);
    ColTitre.Add (formatDateTime('mmmyy',encours));
    LesMois := LesMois+champ+';' ;
    LesColonnes := LesColonnes+';'+ champ;
    TOBLigne.addchampsupvaleur (champ,0,false);
    encours := DebutDeMois(PlusMois (encours,1));
  end;
  if (periodechange) then PGIBoxAf(TexteMsgAffaire[7],Titre);
end;

procedure TOF_AFBUDGET.ChargeGridBudget ;
var TOBLAF,TOBLGD :TOB;
    i_ind,numero :integer;
    Mtcumul : double;
begin
  numero := 0;MtCumul := 0; TOBLGD:=nil;
  for i_ind := 0 to TobAfBudget.detail.count-1 do
  begin
    TOBLAF := TobAfBudget.detail[i_ind];
    if (numero<>TOBLAF.getvalue('ABU_NUMERO')) then
    begin
      MtCumul := 0;
      numero := TOBLAF.getvalue('ABU_NUMERO');
      TOBLGD := tob.create ('Ligne budget',TOBGdBudget,-1);
      TOBLGD.Dupliquer (TOBLigne, False, True, True) ;
      TOBLGD.putvalue('TYPE',TOBLAF.getvalue('ABU_TYPEARTICLE'));
      TOBLGD.putvalue('ARTICLE',TOBLAF.getvalue('ABU_ARTICLE'));
      TOBLGD.putvalue('CODE',TOBLAF.getvalue('ABU_CODEARTICLE'));
      TOBLGD.putvalue('GENE',TOBLAF.getvalue('ABU_ALIMBUDECH'));
    end;
    TOBLGD.putvalue( formatDateTime('mm/yy',TOBLAF.getvalue('ABU_DATEBUD')) ,TOBLAF.getvalue('ABU_MTPVBUDDEV'));
    MtCumul := MtCumul + TOBLAF.getvalue('ABU_MTPVBUDDEV');
    TOBLGD.putvalue('CUMUL',MtCumul);
  end;
end;

////////////// Ligne des Totaux  ////////////////////////////////////
procedure TOF_AFBUDGET.AffecteCumul ;
var     Coord : TRect;
begin
  if  CUMBUDGET<>nil then exit;
  CUMBUDGET:=THNumEdit.Create(TFVierge(Ecran));
  CUMBUDGET.Parent:=PSAICUMUL;
  CUMBUDGET.ParentColor:=True;
  CUMBUDGET.Font.Style:=PSAICUMUL.Font.Style;
  CUMBUDGET.Font.Size:=PSAICUMUL.Font.Size;
  CUMBUDGET.Masks.PositiveMask:=GS.ColFormats[C_cumul];
  CUMBUDGET.Ctl3D:=False; CUMBUDGET.Top:=1;
  Coord:=GS.CellRect(C_cumul,0);
  CUMBUDGET.Left:=Coord.Left + 1;
  CUMBUDGET.Width:=GS.ColWidths[C_cumul] + 1;
  CUMBUDGET.Height:=PSAICUMUL.Height;
  CUMBUDGET.Value:=TOBGdBudget.Somme('CUMUL',[''],[''],False);
end;

procedure TOF_AFBUDGET.OnUpdate ;
var i_ind :integer;
    io:TIOErr;
begin
  Inherited ;
  GS := THGRID(GetControl('GRIDBUDG'));
  if (GS.Col=C_type) or( GS.Col=C_Generation) then
    PostMessage(GS.ValCombo.Handle, WM_KEYDOWN, VK_TAB,  0);
  Application.processMessages;
  TobAfBudget.cleardetail;
  TypeBudg := THValComboBox(GetControl('ABU_TYPEAFBUDGET')).value;
  ////////////// Suppression des lignes sans type ou code articles /////////////////
  for i_ind :=GS.RowCount-1 downto GS.fixedrows do
  begin
    if (trim (GS.cells[C_type,i_ind])='') and (GS.cells[C_Article,i_ind]='')  then
    SupprimeLigne(i_ind);
  end;
  i_ind := GS.fixedrows;
  if (trim (GS.cells[C_type,i_ind])='') and (GS.cells[C_Article,i_ind]='') then
     GS.DeleteRow(i_ind);

  if GS.RowCount > 1 then traitebudget;
  if Erreur then exit;
  io:=Transactions(ValideBudget,1) ;
  Case io of
     oeUnknown : begin
                 PGIBoxAf(TexteMsgAffaire[4], Titre);
                 Exit;
                 end ;
      oeSaisie : begin
                 PGIBoxAf(TexteMsgAffaire[5], Titre);
                 Exit;
                 end ;
     end ;

END;

  ////////////// Sélection des lignes budgets de l'affaires /////////////////
function TOF_AFBUDGET.ChargeBudget (Tobbudg:TOB):boolean ;
var Q : TQuery;
    stSQL : string;
begin
  stSQL := 'SELECT * FROM AFBUDGET WHERE ABU_TYPEAFBUDGET="'+TypeBudg+'" AND ABU_AFFAIRE="'+StAffaire+'"';
  stSQL := stSQL + ' ORDER BY ABU_NUMERO ,ABU_DATEBUD';
  Q:= OpenSQL(stSQL,True);
  result := Not Q.EOF;
  if Not Q.EOF then Tobbudg.LoadDetailDB ('AFBUDGET','','',Q,False,true);
  Ferme(Q);
end;

procedure TOF_AFBUDGET.TraiteBudget ;
var i_ind :integer;
    NumBudget :integer;
    iJour :word;
    dateencours :TDateTime;
    ChampMois,ListeMois,CodeArtUni,CodeArt,stAlim:string;
    MtDev,MtPivot : double;
    TobDetAct :TOB;
    QQ : TQuery;
begin
  for i_ind :=GS.fixedrows to GS.RowCount-1 do
  begin
    TOBL := TOB(GS.Objects[C_Objects,i_ind]);
    if (TOBL = nil) then continue;
    TOBL.InitValeurs;
    TOBL.GetLigneGrid(GS,i_ind,'FIXED;'+LesColonnes );
    codeart := trim(TOBL.getvalue('CODE'));
    if (codeart = '') then continue;
  ////////////// Controle validité du code article /////////////////    
    CodeArtUni := CodeArticleUnique(codeart,'','','','','');
    QQ:=OpenSQL('SELECT GA_TYPEARTICLE,GA_CODEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+ CodeArtUni +'"',true);
    if not QQ.EOF then
    begin
      TOBL.putvalue('TYPE',QQ.findField('GA_TYPEARTICLE').asString) ;
      TOBL.putvalue('CODE',QQ.findField('GA_CODEARTICLE').asString) ;
      TOBL.putvalue('ARTICLE',CodeArtUni);
      ferme(QQ);
      TOBL.PutLigneGrid(GS, i_ind, false, false, LesColonnes);
    end
    else begin
      ferme(QQ);
      PGIBoxAf(TexteMsgAffaire[8],Titre);
      GS.Col := C_article; GS.Row := i_ind; Erreur:=true;
      Exit;
    end;
  end;

  iJour:=1; NumBudget:=0;
  TOBGdBudget.GetGridDetail(GS,GS.rowcount-1,'SaisieBudget','FIXED;'+LesColonnes);
  ////////////// Crée une ligne dans la table pour chaque mois  /////////////////
  for i_ind := 0 to TOBGdBudget.detail.count-1 do
  begin
    inc(NumBudget);ListeMois := LesMois;
    ChampMois:=(Trim(ReadTokenSt(ListeMois)));
    While (ChampMois <> '') do
    begin
      dateencours := encodedate(strtoint(copy(champmois,4,2))+2000,strtoint(copy(champmois,1,2)),ijour);
      TobDetAct := TOB.Create ('AFBUDGET', TobAfBudget,-1);
      TobDetAct.InitValeurs;
      TobDetAct.PutValue('ABU_TYPEAFBUDGET', TypeBudg);
      TobDetAct.PutValue('ABU_TIERS'   ,     TobAff.GetValue('AFF_TIERS'));
      TobDetAct.PutValue('ABU_AFFAIRE' ,     TobAff.GetValue('AFF_AFFAIRE'));
      TobDetAct.PutValue('ABU_AFFAIRE0',     TobAff.GetValue('AFF_AFFAIRE0'));
      TobDetAct.PutValue('ABU_AFFAIRE1',     TobAff.GetValue('AFF_AFFAIRE1'));
      TobDetAct.PutValue('ABU_AFFAIRE2',     TobAff.GetValue('AFF_AFFAIRE2'));
      TobDetAct.PutValue('ABU_AFFAIRE3',     TobAff.GetValue('AFF_AFFAIRE3'));
      TobDetAct.PutValue('ABU_AVENANT' ,     TobAff.GetValue('AFF_AVENANT'));
      TobDetAct.putvalue('ABU_NUMERO',NumBudget);
      TobDetAct.putvalue('ABU_DATEBUD',dateencours);
      TobDetAct.putvalue('ABU_PERIODE',GetPeriode(dateencours));   //mcd 08/08/02
      TobDetAct.putvalue('ABU_SEMAINE',NumSemaine(dateencours));   // mcd 08/08/02
      TobDetAct.putvalue('ABU_TYPEARTICLE',TOBGdBudget.detail[i_ind].getvalue('TYPE'));
      TobDetAct.putvalue('ABU_ARTICLE',TOBGdBudget.detail[i_ind].getvalue('ARTICLE'));
      TobDetAct.putvalue('ABU_CODEARTICLE',TOBGdBudget.detail[i_ind].getvalue('CODE'));
      stAlim := TOBGdBudget.detail[i_ind].getvalue('GENE');
      if (trim(stAlim)='') then stAlim := 'EC1';
      TobDetAct.putvalue('ABU_ALIMBUDECH',stAlim);
      TobDetAct.putvalue('ABU_MTPVBUDDEV',TOBGdBudget.detail[i_ind].getvalue(Champmois));
//    MtDev:= StrToFloat( TOBGdBudget.detail[i_ind].getvalue(Champmois));
      MtDev:= valeur( TOBGdBudget.detail[i_ind].getvalue(Champmois));  // AB-2002-09-13
      // C.B 18/06/2003 Suppression contrevaleur
      ConvertDevToPivot(Devise,MtDev,MtPivot);
      TobDetAct.PutValue('ABU_MTPVBUD',Mtpivot);
      ChampMois:=(Trim(ReadTokenSt(ListeMois)));
    end;
  end;
                 
end;                            

  ////////////// Annule et remplace le budget saisie /////////////////
procedure TOF_AFBUDGET.ValideBudget ;
var TOBDelete :TOB;
begin
  TOBDelete := TOB.Create('maj', nil, -1);
  if ChargeBudget(TOBDelete) then
    TOBDelete.DeleteDB (False);
  TOBDelete.free;
  if (TobAfBudget.Detail.count > 0) then
    TobAfBudget.InsertDB(Nil);
end;

  ////////////// Suppression du budget de l'affaire/////////////////
procedure TOF_AFBUDGET.OnDelete ;
var TOBDelete :TOB;
begin
  Inherited ;
  if (PGIAskAF (TexteMsgAffaire[6], Titre)=mrYes) then
  begin
    TOBDelete := TOB.Create('maj', nil, -1);
    if ChargeBudget(TOBDelete) then
      TOBDelete.DeleteDB (False);
    TOBDelete.free;
    TobAfBudget.cleardetail;
    TOBGdBudget.cleardetail;
    OnLoad ;
  end;
end ;

procedure TOF_AFBUDGET.OnClose ;
begin
  Inherited ;
If Erreur then begin
   LastError:=(-1);
   LastErrorMsg:='';
   Erreur:=false;
   exit;     // il y eu une erreur, on reste sur la même fiche
   end;
  FermeTout;
end ;

procedure TOF_AFBUDGET.OuvreTout ;
BEGIN
  TOBAff:=TOB.Create('AFFAIRE',Nil,-1);
  TOBLigne := Tob.create ('Modele ligne',nil,-1);
  TOBGdBudget:=tob.create('Saisie budget',Nil,-1);
  TobAfBudget := Tob.create('Table Budget', nil, -1);
END;

procedure TOF_AFBUDGET.FermeTout ;
BEGIN
  TobAff.free;  TobAff := NIL;
  TOBLigne.free; TOBLigne:= NIL;
  CUMBUDGET.Free; CUMBUDGET := NIL;
  TOBGdBudget.cleardetail; TOBGdBudget.free; TOBGdBudget:= NIL;
  TobAfBudget.cleardetail; TobAfBudget.free; TobAfBudget:= NIL;
END;

procedure TOF_AFBUDGET.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action=taConsult then Exit ;
  GS.ElipsisButton:=false;
  if not Cancel and (GS.Col=C_Article) then
  begin
    GS.ElipsisButton:=true; GS.ElipsisHint:=TexteMsgAffaire[1] ;
  end;
end;

procedure TOF_AFBUDGET.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var i_ind :integer;
    Cumul :double;
begin
  if Action=taConsult then Exit ;
  if (ACol=C_type) then
    TOB(GS.objects[C_objects,ARow]).putvalue('TYPE',copy (GS.CellValues[ACol, ARow],1,3))
  else if (ACol>=C_mois) then
  begin
    Cumul := 0;      // calcul la colonne Mt cumul
    for i_ind:=C_mois to C_mois+NombreMois-1 do
      Cumul:=Cumul + Valeur(GS.cells[i_ind,ARow]);
    GS.cells[C_cumul,ARow] := StrF00(Cumul,0);
    TOB(GS.objects[C_objects,ARow]).putvalue('CUMUL', Cumul);
    CUMBUDGET.Value:=TOBGdBudget.Somme('CUMUL',[''],[''],False);
  end
  else if  (ACol=C_Cumul) then
  begin
    TOB(GS.objects[C_objects,ARow]).putvalue('CUMUL', Valeur(GS.cells[C_Cumul,ARow]));
    CUMBUDGET.Value:=TOBGdBudget.Somme('CUMUL',[''],[''],False);
  end;
end;

  ////////////// Recherche code article /////////////////
procedure TOF_AFBUDGET.GSElipsisClick(Sender: TObject);
Var ARTICLE : THCritMaskEdit;
    Coord : TRect;
BEGIN
  Coord := GS.CellRect (GS.Col, GS.Row);
  ARTICLE := THCritMaskEdit.Create (ECRAN);
  ARTICLE.Parent := GS;
  ARTICLE.Top := Coord.Top;
  ARTICLE.Left := Coord.Left;
  ARTICLE.Width := 3; ARTICLE.Visible := False;
  ARTICLE.Text:= GS.Cells[GS.Col,GS.Row] ;
  ARTICLE.DataType:='GCARTICLEGENERIQUE';
  DispatchRecherche(ARTICLE,1,'GA_TYPEARTICLE='+TOB(GS.objects[C_objects,GS.Row]).getvalue('TYPE'),
  'GA_CODEARTICLE='+ARTICLE.text,'');
  if ARTICLE.Text <> '' then
  begin
    GS.Cells[GS.Col,GS.Row]:= Trim (Copy (ARTICLE.Text, 1, 18));
    TOB(GS.objects[C_objects,GS.Row]).putvalue('ARTICLE',ARTICLE.Text);
  end;
  ARTICLE.Destroy;
end;

procedure TOF_AFBUDGET.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  GS.InvalidateRow(ou) ;
end;

procedure TOF_AFBUDGET.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  GS.InvalidateRow(ou) ;
end;

procedure TOF_AFBUDGET.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
begin
FocusGrid := False;  ARow:=0;
if((Screen.ActiveControl = GS) or (Screen.ActiveControl.Owner = GS)) then
    BEGIN
    FocusGrid := True;
    ARow := GS.Row;
    END ;
Case Key of
    VK_INSERT : BEGIN
        if FocusGrid then
            BEGIN
            Key := 0;
            InsertLigne (ARow);
            END;
        END;
    VK_DOWN   : BEGIN
        if FocusGrid and (ARow=GS.RowCount-1) then
            BEGIN
            Key := 0;
            InsertLigne (ARow);
            END;
                END;
    VK_DELETE : BEGIN
        if (FocusGrid) and (Shift=[ssCtrl]) then
            BEGIN
            Key := 0 ;
            SupprimeLigne (ARow) ;
            END ;
        END;
    END;
end;

procedure TOF_AFBUDGET.InsertLigne(var ARow:integer) ;
begin
  PostMessage(GS.ValCombo.Handle, WM_KEYDOWN, VK_TAB,  0) ;
  Application.processMessages;
  GS.CacheEdit; GS.SynEnabled := false;
  inc (Arow); GS.InsertRow (ARow);
  GS.Row := ARow;GS.Rows[ARow].Clear;
  GS.MontreEdit; GS.SynEnabled := true;
  TOBL := Tob.create ('Ligne budget',TOBGdBudget,ARow-1);
  TOBL.Dupliquer (TOBLigne, False, True, True) ;
  TOBL.PutLigneGrid(GS, ARow, false, false, LesColonnes);
  GS.ShowCombo(C_type,GS.Row) ;
end;

procedure TOF_AFBUDGET.SupprimeLigne(ARow:integer) ;
begin
  if (GS.RowCount = 2) then
  begin
    TOB(GS.objects[0,ARow]).Dupliquer (TOBLigne, False, True, True) ;
    TOB(GS.objects[0,ARow]).PutLigneGrid(GS, ARow, false, false, 'FIXED;'+LesColonnes);
    GS.ShowCombo(C_type,ARow) ;GS.CellValues[c_type,Arow] := '' ;
  end;
  if (GS.RowCount < 3) then Exit ;
  GS.CacheEdit; GS.SynEnabled := False;
  if (GS.objects[C_Objects,ARow]<>Nil) then TOB(GS.objects[0,ARow]).Free;
  GS.DeleteRow(ARow) ;
  GS.MontreEdit; GS.SynEnabled := True;
end ;

procedure TOF_AFBUDGET.GSComboPopUp(Sender: TObject);
begin
  if (GS.Col =C_type) then  GS.valcombo.plus := PlusTypeArticle ;
end;

procedure TOF_AFBUDGET.BDELLINEClick(Sender: TObject);
var Arow :longint;
begin
  if Action=taConsult then Exit ;
  Arow := GS.row; SupprimeLigne(Arow);
end;

procedure TOF_AFBUDGET.BNEWLINEClick(Sender: TObject);
var Arow :longint;
begin
  if Action=taConsult then Exit ;
  Arow := GS.row; InsertLigne(Arow);
end;

  ////////////// Répartition sur les mois du cumul saisie /////////////////
procedure TOF_AFBUDGET.BCUMULClick(Sender: TObject);
var Arow :longint;
    MtCumul : double;
    i_ind,iMtMois,iMtDernier : integer;
begin
  Arow := GS.row;  MtCumul := Valeur(GS.cells[C_cumul,ARow]);
  if (Action=taConsult) or (MtCumul=0) then Exit ;
  iMtMois := trunc(MtCumul / NombreMois);
  iMtDernier := trunc(MtCumul) - (iMtMois*(NombreMois-1));
  for i_ind:= C_mois to C_mois+NombreMois-2 do
    GS.Cells[i_ind,ARow] := intToStr(iMtMois);
  GS.Cells[C_mois+NombreMois-1,ARow] := intToStr (iMtDernier);
end;

  ////////////// Repositionne le montant sur la ligne des totaux /////////////////
procedure TOF_AFBUDGET.GSDessineCell(ACol, ARow: Integer; Canvas: TCanvas;
  AState: TGridDrawState);
  var Coord : TRect;
begin
  If Arow < GS.Fixedrows then exit ;
  if  (CUMBUDGET<>nil) and ((GS.ScrollBars=ssBoth) or (GS.ScrollBars=ssHorizontal)) then
  begin
      Coord:=GS.CellRect(C_cumul,0);
      CUMBUDGET.Left:=Coord.Left + 1;
  end;
end;

Initialization
  registerclasses ( [ TOF_AFBUDGET ] ) ;
end.
