Unit UTofAfProfilGener;                                
                    
interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, M3FP ,
      HCtrls,HEnt1,HMsgBox,UTOF, UTob,Grids,EntGC,HTB97,paramsoc,
{$IFDEF EAGLCLIENT}
      eTablette,Maineagl,
{$ELSE}
      Tablette, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,  
{$ENDIF}
      Dicobtp,graphics,Saisutil,
      UtilPaieAffaire,menus,utilarticle,utot,vierge;
                                      
Type
     TOF_afprofilgener = Class (TOF)
     public
        Faffiche,FDispo : TListBox;
        zmodif : boolean;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure OnDelete ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose;  override ;

        procedure RecupLibAcpte(valeur ,zone : string);
     private
        TOBAffich,TOBDispo : TOB;
        Profilcharge : string;
        procedure TestEvent (Sender: TObject);
        procedure TestEvent2 (Sender: TObject);
        procedure ChargeStandard (Sender: TObject);
        Procedure ChargeLesTob(CodeProfil : string);
        procedure ClearEcran;
        Procedure ChargeLesListBox;
        Procedure ToutAllouer;
        Procedure ToutLiberer;
        procedure BEnleverClick(Sender: TObject);
        procedure BAjouterClick(Sender: TObject);
        procedure BMonterClick(Sender: TObject);
        procedure BDescendreClick(Sender: TObject);
        procedure FAfficheClick(Sender: TObject);
        procedure Init_Liste(Sender : TObject);
        procedure AfficheEltGen(Tobdet : TOB);
        procedure AskMajProfil;
        procedure MajProfil(Code : string);
        function  CtrlValidite : boolean;
        procedure TrtReplibech;
        procedure TrtDetail; 
        procedure TrtCumul;
        procedure InitZones;
    END ;

    TOT_AFPROFILGENER = Class (TOT)
    procedure OnDeleteRecord  ; override ;
    procedure OnLoad ; virtual ;
    End;


Type TListeCol = Class
  element : String;
  lib : string;
  typearticle : string;
  article,codearticle,Replibech,Detail : string;
  cumul : boolean;
  libart,libsup : string;
  seuilmaxi,seuilmini : double;
END;

     const
	// libellés des messages de la TOF  afprofilgener
	TexteMsgAffaire: array[1..2] of string 	= (
          {1}        'Cet élément n''est pas compatible avec les précédents'
          {2}        ,'Code Prestation invalie'
                             );
   var titre : string;
Procedure AFLanceFiche_ProfilGener(Argument:string);

implementation

{ TOF_afprofilgener}

procedure TOF_AFPROFILGENER.OnLoad;
Begin
 Inherited;
  setcontroltext('APG_PROFILGENER',GetParamSoc('SO_AFPROFILGENER'));
  titre := ecran.caption;
End;

procedure TOF_AFPROFILGENER.OnClose;
Begin
 AskMajProfil;
 Toutliberer;
End;

procedure TOF_AFPROFILGENER.OnArgument(stArgument : String );
Var CC1,CC2,CC3,CC4,CC5 : THedit;

 Combo   : THValComboBox;
 BtAjouter, BtEnlever,BtDescendre,BtMonter : TToolBarButton97;
BEGIN
                      
Inherited;
// ruse pour ne pas passer par le script
CC1 := THEDIT(GetControl('APG_SEUILMINI'));
CC1.OnExit:=TestEvent2;
CC2 := THEDIT(GetControl('APG_SEUILMAXI'));
CC2.OnExit :=TestEvent2;
CC3 := THEDIT(GetControl('APG_LIBART'));
CC3.OnExit :=TestEvent2;
CC4 := THEDIT(GetControl('APG_CODEARTICLE'));
CC4.OnExit :=TestEvent2;
CC5 := THEDIT(GetControl('APG_LIBSUP'));
CC5.OnExit :=TestEvent2;


Combo   := THValComboBox   (GetControl('APG_PROFILGENER'));
Combo.OnChange:= ChargeStandard;

BtAjouter := TToolBarButton97 (GetControl('BAJOUTER')); BtAjouter.OnClick := BAjouterClick;
BtEnlever := TToolBarButton97 (GetControl('BENLEVER')); BtEnlever.OnClick := BEnleverClick;
BtDescendre := TToolBarButton97 (GetControl('BDESCENDRE')); BtDescendre.OnClick := BDescendreClick;
BtMonter := TToolBarButton97 (GetControl('BMONTER')); BtMonter.OnClick := BMonterClick;

Faffiche   := TListBox   (GetControl('FAFFICHE'));
FDispo   := TListBox   (GetControl('FDISPO'));

Faffiche.OnClick:=TestEvent;
Faffiche.OnDblClick:=BEnleverClick;
Fdispo.OnDblClick:=BAjouterClick;


InitZones;

END;

procedure TOF_AFPROFILGENER.InitZones;
BEGIN
SetControlProperty('APG_LIBART','Color',clbtnFace);
SetControlEnabled('APG_LIBART',false);
SetControlProperty('APG_LIBSUP','Color',clbtnFace);
SetControlEnabled('APG_LIBSUP',false);
SetControlProperty('APG_CODEARTICLE','Color',clbtnFace);
SetControlEnabled('APG_CODEARTICLE',false);
SetControlProperty('TAPG_CODEARTICLE','Color',clbtnFace);
SetControlEnabled('TAPG_CODEARTICLE',false);
SetControlEnabled('APG_SEUILMINI',false);
SetControlProperty('APG_SEUILMINI','Color',clbtnFace);
SetControlEnabled('APG_SEUILMAXI',false);
SetControlProperty('APG_SEUILMAXI','Color',clbtnFace);
SetControlEnabled('TAPG_SEUILMINI',false);
SetControlEnabled('TAPG_SEUILMAXI',false);
SetControlEnabled('APG_REPLIBECH',false);
SetControlEnabled('APG_DETAIL',false);
END;

procedure TOF_AFPROFILGENER.OnUpdate;
Var Code: string;
BEGIN
    Code := GetControlText('APG_PROFILGENER');
    MajProfil(Code);
    tobaffich := NIL;
    zmodif := false;
END;

procedure TOF_AFPROFILGENER.AskMajProfil;
var lib,stit : string;
BEGIN
// if (tobaffich <> NIL) then
    if (zmodif)  then
    Begin
      Lib :=  RechDom('AFPROFILGENER',ProfilCharge,False);
      stit := format('Profil facture : %s - %s',[ProfilCharge,Lib]);
      If (PGIAskAF('Voulez vous enregistrer ce profil ?',stit)= mrYes) then
      Begin
        MajProfil(ProfilCharge);
      End;
      tobaffich := NIL;
    End;
END;

procedure TOF_AFPROFILGENER.MajProfil(Code : string);
Var Req : string;
    TobMaj,TobDetMaj : TOB;
    Cpt,wi : Integer;
    C : TlisteCol;
    zelt2 : string;
BEGIN
    // supression des elements existants
    Req := 'DELETE FROM PROFILGENER WHERE APG_PROFILGENER = "'+code+'"';
    ExecuteSQL(Req) ;

    // création d'une Tob avec les enreg à créer.
    TobMaj := TOB.create('Creat Profil',NIL,-1);

    // création enreg A00 (Multi affaire)
    Cpt := 0;
    TobDetMaj := Tob.create('PROFILGENER',TobMaj,-1);
    TobDetMaj.PutValue('APG_PROFILGENER',code);
    TobDetMaj.PutValue('APG_NUMORDRE',Cpt);
    TobDetMaj.PutValue('APG_ELEMENT','A00');

    if  (TCheckBox(GetControl('APG_CUMUL')).checked) Then
       TobDetMaj.PutValue('APG_CUMUL', 'X')
    else
        TobDetMaj.PutValue('APG_CUMUL', '-');

    // création des autres enregs
    For wi:=0 to FAffiche.Items.Count-1 do
    begin
      Inc(Cpt);
      C:=TListeCol(FAffiche.Items.Objects[wi]) ;
      TobDetMaj := Tob.create('PROFILGENER',TobMaj,-1);
      TobDetMaj.PutValue('APG_PROFILGENER',code);
      TobDetMaj.PutValue('APG_NUMORDRE',Cpt);
      TobDetMaj.PutValue('APG_ELEMENT',C.element);
      TobDetMaj.PutValue('APG_CODEARTICLE',C.codearticle);
      TobDetMaj.PutValue('APG_ARTICLE',C.article);

      TobDetMaj.PutValue('APG_LIBART',C.libart);
      TobDetMaj.PutValue('APG_LIBSUP',C.libsup);
      TobDetMaj.PutValue('APG_SEUILMINI',C.seuilmini);
      TobDetMaj.PutValue('APG_SEUILMAXI',C.seuilmaxi);

      if (C.element = 'G00') then
        TobDetMaj.PutValue('APG_REPLIBECH',C.Replibech)
      else
         TobDetMaj.PutValue('APG_REPLIBECH','-');

      zelt2 := copy(C.element,2,1);
      if (zelt2 = '2')  or (zelt2  = '3')then
        TobDetMaj.PutValue('APG_DETAIL',C.Detail)
      else
         TobDetMaj.PutValue('APG_DETAIL','-');

    End;

    TOBMaj.InsertOrUpdateDB(False) ;
    TobMaj.free;

END;

procedure TOF_AFPROFILGENER.OnDelete;
Var Code , stit,lib,Req: string;
  res : boolean;
   Combo   : THValComboBox;
BEGIN
  Code := GetControlText('APG_PROFILGENER');
  Lib :=  RechDom('AFPROFILGENER',Code,False);
  stit := format('Profil facture : %s - %s',[Code,Lib]);
  If ( Code = GetParamSoc('SO_AFPROFILGENER')) then
    Begin
      PGIBoxAF('La suppression de ce profil est interdite',stit);
    End
  else
   Begin
      // controle des liens avec  la table Affaire
      res := SupTablesLiees ('AFFAIRE', 'AFF_PROFILGENER', Code, '' , false);
      if (res) then
        PGIBoxAF('Suppression interdite, ce profil est utilisé',stit)
      else
        Begin
          If (PGIAskAF('Confirmez-vous la suppression de ce profil ?',stit)= mrYes) then
            Begin
            // Suppresion du profil
            SupTablesLiees ('PROFILGENER', 'APG_PROFILGENER', Code, '' , true);
            // Suppression de la tablette des elements possible (AFPROFILGENERELT)
            Req :='DELETE FROM CHOIXCOD WHERE CC_TYPE="APG" and CC_CODE="'+Code+'"';
            Executesql(Req);

            Combo := THValComboBox (GetControl('APG_PROFILGENER'));
            AvertirTable('AFPROFILGENER');    // Rechargement de la tablette
            Combo.reload; Combo.Value:= GetParamSoc('SO_AFPROFILGENER');
            End;
        End;
   End;                                                    
END;

procedure TOF_AFPROFILGENER.TrtCumul ;
Var C1 : TListeCol ;
    wi : integer;
    C1elt : string;  // éléments dèjà sélectionnés
BEGIN
zmodif := true;
// SI on demande une présentation cumulée des affaires,
// on ne peut pas avoir  de présentataion par sous-affaire
if  (TCheckBox(GetControl('APG_CUMUL')).checked) Then
Begin
   For wi:=0 to FAffiche.Items.Count-1 do
    begin
      C1:=TListeCol(FAffiche.Items.Objects[wi]) ;
      C1elt  :=  C1.element;

      if ((C1elt = 'AST') or (C1elt = 'ACO'))  then
        Begin
          PGIInfoAf(TexteMsgAffaire[1],titre);
          exit;
        End;
    End; //for
End;

END;

procedure TOF_AFPROFILGENER.TrtReplibech ;
Var Box : TCheckBox;
begin
zmodif := true;
  Box := TCheckBox(GetControl('APG_REPLIBECH'));
  If (Box.Checked) then
    Begin
    TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).replibech := 'X' ;
    TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).libart := '';
    SetControlText('APG_LIBART','');
    SetControlProperty('APG_LIBART','Color',clbtnFace);
    SetControlEnabled('APG_LIBART',false);
    End
  else
    Begin
    TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).replibech := '-';
    SetControlProperty('APG_LIBART','Color',clWindow);
    SetControlEnabled('APG_LIBART',true);
    SetFocusControl('APG_LIBART');
    THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_LIBACOMPTE'));
    TPopUpMenu(GetControl('POP_LIBACOMPTE')).autopopup := true;
    End;
END;

procedure TOF_AFPROFILGENER.TrtDetail ;
Var Box : TCheckBox;
begin
zmodif := true;
  Box := TCheckBox(GetControl('APG_DETAIL'));
  If (Box.Checked) then
    Begin
    TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).detail := 'X' ;
    End
  else
    Begin
    TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).detail := '-';
  End;
END;

procedure TOF_AFPROFILGENER.TestEvent (Sender: TObject);
begin
FAfficheClick(nil) ;
end;

procedure TOF_AFPROFILGENER.TestEvent2 (Sender: TObject);
var  zmini,zmaxi : double;
    st, dim : string;
BEGIN

if ((Sender is THEdit) and (Thedit(Sender).Name='APG_SEUILMINI') and  (Thedit(Sender).Modified)) then
  begin
    zmodif := true;
    Thedit(Sender).Modified :=False ;
    st := Thedit(Sender).text;
    zmini := 0;
    if (st <> '') then
      Begin
      zmini := STRToFloat(st);
      End;

    if (FAffiche.ItemIndex>=0)  then
      Begin
      TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).seuilmini := zmini;
      End;
  End;

if ((Sender is THEdit) and (Thedit(Sender).Name='APG_SEUILMAXI') and  (Thedit(Sender).Modified)) then
  begin
    zmodif := true;
    Thedit(Sender).Modified :=False ;
    st := Thedit(Sender).text;
    zmaxi:= 0;
    if (st <> '') then
      Begin
      zmaxi := STRToFloat(st);
      End;

    if (FAffiche.ItemIndex>=0)  then
      Begin
      TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).seuilmaxi:= zmaxi;
      End;
  End;


if ((Sender is THEdit) and (Thedit(Sender).Name='APG_LIBART') and  (Thedit(Sender).Modified)) then
  begin
    zmodif := true;
    Thedit(Sender).Modified :=False ;
    st := Thedit(Sender).text;

    if (FAffiche.ItemIndex>=0)  then
      Begin
      TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).libart:= st;
      End;
  End;

  if ((Sender is THEdit) and (Thedit(Sender).Name='APG_LIBSUP') and  (Thedit(Sender).Modified)) then
  begin
    zmodif := true;
    Thedit(Sender).Modified :=False ;
    st := Thedit(Sender).text;

    if (FAffiche.ItemIndex>=0)  then
      Begin
      TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).libsup:= st;
      End;
  End;
   // mcd 10/03/03
if ((Sender is THEdit) and (Thedit(Sender).Name='APG_CODEARTICLE')
  and (FAffiche.ItemIndex>=0)
  and  (Thedit(Sender).text <>TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).Codearticle)) then   // on regarde si code article à changer .. car si on lcic sur ... thedit.modified n'est pas changer
  begin
    st := Thedit(Sender).text;
    if Not(ExisteSql('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+st+'"')) then
      begin
      SetFocusControl ('APG_CODEARTICLE');
      exit;
      end;
    zmodif := true;
    Thedit(Sender).Modified :=False ;

    if (FAffiche.ItemIndex>=0)  then
      Begin
      TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).Codearticle:= st;
      dim :='';
      TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]).article:= CodeArticleUnique(st,dim,dim,dim,dim,dim);
      End;
  End;

END;


procedure TOF_AFPROFILGENER.ChargeStandard (Sender: TObject);
Var Combo   : THValComboBox;
    CodeProfil: string;
BEGIN
initzones;
// Récupération des composants utilisés
Combo := THValComboBox (GetControl ('APG_PROFILGENER') );

if (Combo.Text = '')  then
    PgiBoxAF('Vous devez saisir un code profil par défaut dans vos paramètres, gestion affaire, préférence affaire',Titre)
else
    Begin
    AskMajProfil;
    CodeProfil := Combo.value;
    ProfilCharge := Combo.value;
    Toutliberer;
    Toutallouer;
    ClearEcran;
    ChargeLesTob(CodeProfil);
    ChargeLesListBox;
    zmodif := false;
    End;
END;

procedure TOF_AFPROFILGENER.ToutAllouer;
BEGIN
TOBAffich:=TOB.Create('Profil Affiché',Nil,-1) ;
TOBDispo:=TOB.Create('Profil Dispo',Nil,-1) ;
END;

procedure TOF_AFPROFILGENER.Toutliberer;
var i : integer;
BEGIN
TOBAffich.Free ; TOBAffich:=Nil ;
TOBDispo.Free ; TOBDispo:=Nil ;

for i:=0 to FAffiche.Items.Count-1 do TListeCol(FAffiche.Items.Objects[i]).Free ;
for i:=0 to FDispo.Items.Count-1 do TListeCol(FDispo.Items.Objects[i]).Free ;
Faffiche.Items.Clear ; FDispo.Items.Clear;

END;

procedure TOF_AFPROFILGENER.ClearEcran;
BEGIN
  Fdispo.clear;
  Faffiche.clear;

  SetControlText('APG_ARTICLE','');
  SetControlText('APG_CODEARTICLE','');
  SetControlChecked ('APG_CUMUL',false);
  SetControlChecked ('APG_REPLIBECH',false);
  SetControlChecked ('APG_DETAIL',false);
  SetControlText('APG_LIBART','');
  SetControlText('APG_LIBSUP','');
  SetControlText('APG_SEUILMINI','');
  SetControlText('APG_SEUILMAXI','');

END;

procedure TOF_AFPROFILGENER.ChargeLesListBox;
Var  wi : integer;
      TobDispoDet,TobAffichDet: Tob;
      Cdispo : TListeCol;
      Caffi : TListeCol;
      zelt,zelt1 :string;
BEGIN
  for wi:=0  to TobDispo.Detail.count-1 do
  Begin  // boucle sur les éléments diponibles
      TobDispoDet := TobDispo.Detail[wi];
      zelt := TobDispoDet.GetValue('APG_ELEMENT');
      zelt1 := copy(zelt,1,1);
      if ((GetParamSoc('SO_AFGERESSAFFAIRE') =false) and (zelt1 = 'A')) then
        continue;
      if ((GetParamSoc('SO_AFTRTFTP') =false) and (zelt1 = 'E')) then
        continue;


      Cdispo := TlisteCol.Create;
      Cdispo.element := TobDispoDet.GetValue('APG_ELEMENT');
      Cdispo.lib := TobDispoDet.GetValue('ZLIBELLE');
      FDispo.Items.AddObject(Cdispo.lib,Cdispo);
    End;

  for wi:=0  to TobAffich.Detail.count-1 do
  Begin  // boucle sur les éléments diponibles
    TobAffichDet := TobAffich.Detail[wi];
    zelt := TobAffichDet.GetValue('APG_ELEMENT');
    if (zelt <> 'A00') then
    begin
      Caffi := TlisteCol.Create;
      Caffi.element := TobAffichDet.GetValue('APG_ELEMENT');
      Caffi.lib := TobAffichDet.GetValue('ZLIBELLE');
      Caffi.libart := TobAffichDet.GetValue('APG_LIBART');
      Caffi.libsup := TobAffichDet.GetValue('APG_LIBSUP');
      if (zelt = 'G00') then
      Begin
        Caffi.article := TobAffichDet.GetValue('APG_ARTICLE');
        Caffi.codearticle := TobAffichDet.GetValue('APG_CODEARTICLE');
      End
      else
      Begin
        Caffi.article := '';
        Caffi.codearticle := '';
      End;

      Caffi.seuilmini := TobAffichDet.GetValue('APG_SEUILMINI');
      Caffi.seuilmaxi := TobAffichDet.GetValue('APG_SEUILMAXI');
      Caffi.replibech := TobAffichDet.GetValue('APG_REPLIBECH');
      if (Caffi.replibech = '') then Caffi.replibech := '-';
      Caffi.detail := TobAffichDet.GetValue('APG_DETAIL');
      if (Caffi.detail = '') then Caffi.detail := '-';

      FAffiche.Items.AddObject(Caffi.lib,Caffi);
    end;
  END;
END;

procedure TOF_AFPROFILGENER.ChargeLesTob (CodeProfil : string);
Var QQ : TQuery ;
    Req,wcode,wlib : string;
    nbaff,i : integer;
    Tobdet,TobDetDispo,Tobtrt : TOB;
BEGIN
// Lecture Profil Generation  avec code Profil Saisie
// On doit tout lire puisque mise à jour ensuite... peu d'enrgt et enrgt court
Req := 'SELECT * FROM PROFILGENER WHERE APG_PROFILGENER = "'+codeProfil+'"';
QQ:=OpenSQL(Req,True) ;
If Not QQ.EOF then
  begin
  TOBAffich.LoadDetailDB('PROFILGENER','','',QQ,False) ;
  TOBAffich.AddChampSup('ZLIBELLE',True);
  nbaff := TOBAffich.Detail.count;
  end
else
  nbaff := 0;
Ferme(QQ) ;

if (nbaff <> 0) then
Begin
   TobDet := TobAffich.FindFirst(['APG_ELEMENT'],['A00'],False);
   if TobDet <> NIL Then AfficheEltGen(TobDet);
End;

// Recup des donnees de la tablette des elements possible (AFPROFILGENERELT)
  // mcd 15/05/03 ppr ne pas avoir les lignes CTR en GI
Req := 'SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="APE"';
If ctxScot in V_PGI.PGIContexte then Req :=Req + ' And CO_CODE not like "N%"';
Req := Req +' order by CO_CODE';
QQ:=OpenSQL(Req,True) ;


If QQ.EOF Then exit;
TobTrt := Tob.create ('liste champs', Nil,-1);
TobTrt.LoadDetailDB('','','',QQ,False);
Ferme(QQ) ;

For i := 0 To TobTrt.detail.count-1 do
Begin
   wcode:= TobTrt.detail[i].GetValue('CO_CODE');
   wlib := TraduitGa(TobTrt.detail[i].GetValue('CO_LIBELLE'));   //mcd 12/05/03 ajout traduitga

  if nbaff <> 0 then  // Recherche si l'element est dans la tob Affiche
    Begin
      TobDet := TobAffich.FindFirst(['APG_ELEMENT'],[wcode],False);
    End
  else
    TobDet := NIL;

  If (TobDet = Nil) Then
      BEGIN
      TobDetDispo := Tob.create('PROFILGENER',TobDispo,-1);
      TOBDetDispo.AddChampSup('ZLIBELLE',False);
      TobDetDispo.PutValue('APG_ELEMENT',  wcode);
      TobDetDispo.PutValue('ZLIBELLE', wlib);
      END
  else
      BEGIN
      //Maj du libelle de l'élément dans la tob Affich
      TOBDet.AddChampSup('ZLIBELLE',False);
      TobDet.PutValue('ZLIBELLE',wlib);
      END
  End;
	Tobtrt.free;
END;

// Affichage des élements Généraux  : affaire multi
procedure TOF_AFPROFILGENER.AfficheEltGen(Tobdet : TOB);
BEGIN
   SetControlChecked ('APG_CUMUL',(TobDet.GetValue('APG_CUMUL') = 'X'));
END;

procedure TOF_AFPROFILGENER.BAjouterClick(Sender: TObject);
var oldi : integer ;
    C : TListeCol ;
    ok : boolean;
begin
zmodif := true;
if FDispo.ItemIndex>=0 then
   BEGIN
   OldI:=FDispo.ItemIndex ;
   C:=TListeCol(FDispo.Items.Objects[OldI]) ;
   ok := CtrlValidite;
   if (ok) then
   Begin
   FAffiche.Items.AddObject(FDispo.Items[OldI],C) ;
   FDispo.Items.Delete(OldI) ;
   if OldI<Fdispo.Items.Count then FDispo.itemIndex:=OldI else FDispo.itemIndex:=OldI-1 ;
   FAffiche.ItemIndex:=FAffiche.Items.Count-1 ;
   FAfficheClick(nil) ;
   End;
   END ;
end;

procedure TOF_AFPROFILGENER.BEnleverClick(Sender: TObject);
var oldi : integer ;
    C : TListeCol ;
begin
zmodif := true;
if FAffiche.ItemIndex>=0 then
  BEGIN
  OldI:=FAffiche.ItemIndex ;
  C:=TListeCol(FAffiche.Items.Objects[OldI]) ;
  Init_Liste(C);
  FDispo.Items.AddObject(FAffiche.Items[OldI],C) ;
  FAffiche.Items.Delete(OldI) ;
  if OldI<FAffiche.Items.Count then FAffiche.itemIndex:=OldI else FAffiche.itemIndex:=OldI-1 ;
  FAfficheClick(nil) ;
END ;
                             
END;


procedure TOF_AFPROFILGENER.BMonterClick(Sender: TObject);
var oldi : integer ;
begin
zmodif := true;
OldI:=FAffiche.ItemIndex ;
if OldI>0 then
   BEGIN
   if (FAffiche.Items[OldI-1]='IE_CHRONO') then
     BEGIN
     Exit ;
     END ;
   FAffiche.Items.Exchange(OldI,OldI-1) ;
   FAffiche.ItemIndex:=OldI-1 ;
   END ;
end;

procedure TOF_AFPROFILGENER.BDescendreClick(Sender: TObject);
var oldi : integer ;
begin
zmodif := true;
if (FAffiche.Items[FAffiche.ItemIndex]='IE_CHRONO') then
  BEGIN
  Exit ;
  END ;
OldI:=FAffiche.ItemIndex ;
if ((OldI>=0) and (oldI<FAffiche.items.Count-1)) then
   BEGIN
   FAffiche.Items.Exchange(OldI+1,OldI) ;
   FAffiche.ItemIndex:=OldI+1 ;
   END ;
end;

procedure TOF_AFPROFILGENER.FAfficheClick(Sender: TObject);
Var C : TListeCol ;
    dim,st,zelt2 : String;
BEGIN
c := nil;
if FAffiche.ItemIndex>=0 then
   Begin
    C:=TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]);

    zelt2 := copy(C.element,2,1);
    if ((C.element = 'G00') and (C.article = '')) then
    Begin
         C.codearticle := VH_GC.AfAcompte;
        dim := '';
        st := VH_GC.AFAcompte;
        C.article := CodeArticleUnique(C.codearticle,dim,dim,dim,dim,dim);
    End;

    if (C.element = 'G00') then
    Begin
      SetControlEnabled('APG_REPLIBECH',true);
      if (C.replibech = 'X') then
      Begin
        SetControlChecked ('APG_REPLIBECH',true);
        SetControlText('APG_LIBART','');
        SetControlEnabled('APG_LIBART',false);
        SetControlProperty('APG_LIBART','Color',clWindow);
        C.libart := '';
      End
      else
      Begin
        SetControlChecked ('APG_REPLIBECH',false);
        SetControlEnabled('APG_LIBART',true);
        SetControlProperty('APG_LIBART','Color',clbtnFace);
      End
    End
    else
    Begin
      SetControlEnabled ('APG_REPLIBECH',false);
    End;


    if ((C.element = 'G00') and (C.replibech = '-') and (C.libart = '')) then
    Begin
          C.libart := LibelleArticleGenerique(VH_GC.AfAcompte);
    End;

    SetControlText('APG_LIBART',C.libart);
    SetControlText('APG_LIBSUP',C.libsup);
    SetControlText('APG_ARTICLE',C.article);
    SetControlText('APG_CODEARTICLE',C.codearticle);
    SetControlText('APG_SEUILMINI',StrF00(C.seuilmini,2));
    SetControlText('APG_SEUILMAXI',StrF00(C.seuilmaxi,2));

// Zone APG_CODEARTICLE  , Code acompte en consult ET slt sur G00
    if (C.element = 'G00') then
    Begin
      SetControlEnabled('TAPG_CODEARTICLE',true);
      SetControlEnabled('APG_CODEARTICLE',true);
      SetControlProperty('APG_CODEARTICLE','Color',clWindow);
    End
    else
    Begin
      SetControlEnabled('TAPG_CODEARTICLE',false);
      SetControlEnabled('APG_CODEARTICLE',false);
      SetControlProperty('APG_CODEARTICLE','Color',clbtnFace);
    End;


// Zone APG_LIBART  , libellé ou commentaire affaire
   if ( (C.element = 'P10') or
      (C.element = 'S00') or (C.element = 'S10') or
      (C.element = 'V00') or (C.element = 'V10') or
      (C.element = 'N00') or (C.element = 'N10') or
      (C.element = 'C00') or (C.element = 'TOU') or
      (C.element = 'XPO') or
      ((C.element = 'G00') and (C.replibech = 'X'))) then
      Begin  // pas de libellé

      SetControlProperty('APG_LIBART','Color',clbtnFace);
      SetControlEnabled('APG_LIBART',false);
      End
    else
      Begin
      SetControlProperty('APG_LIBART','Color',clWindow);
      SetControlEnabled('APG_LIBART',true);
      TPopUpMenu(GetControl('POP_COMAFF')).autopopup := false;
      TPopUpMenu(GetControl('POP_LIBACOMPTE')).autopopup := false;
      TPopUpMenu(GetControl('POP_LIBEMP')).autopopup := false;
      TPopUpMenu(GetControl('POP_LIBPRE')).autopopup := false;
      TPopUpMenu(GetControl('POP_LIBSSAFF')).autopopup := false;
      TPopUpMenu(GetControl('POP_LIBSSCLI')).autopopup := false;
      if (C.element = 'P00') then // Libellé prestation
      Begin
        THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_LIBPRE'));
        TPopUpMenu(GetControl('POP_LIBPRE')).autopopup := true;
      End;
      if (C.element = 'C10')  then   // commentaire affaire
      Begin
        THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_COMAFF'));
        TPopUpMenu(GetControl('POP_COMAFF')).autopopup := true;

      End;
      if (C.element = 'G00')  then     // commentaire acompte
      Begin
        THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_LIBACOMPTE'));
        TPopUpMenu(GetControl('POP_LIBACOMPTE')).autopopup := true;
      End;
      // Libellé employé
      if ((C.element = 'ECO') or (C.element = 'EST') or(C.element = 'EPR')) then
      Begin
        THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_LIBEMP'));
        TPopUpMenu(GetControl('POP_LIBEMP')).autopopup := true;
      End;
      // Libellé sous affaire
      if ((C.element = 'ACO') or (C.element = 'AST') ) then
      Begin
        THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_LIBSSAFF'));
        TPopUpMenu(GetControl('POP_LIBSSAFF')).autopopup := true;
      End;
      // Libellé sous total CLient affaire
      if (C.element = 'XCL')  then
      Begin
        THedit(GetControl('APG_LIBART')).PopupMenu := tpopupmenu(GetControl('POP_LIBSSCLI'));
        TPopUpMenu(GetControl('POP_LIBSSCLI')).autopopup := true;
      End;

    End;
// Zone APG_LIBSUP  , libellé ou commentaire affaire
    if (C.element = 'C10')  then   // commentaire affaire
      Begin
        THedit(GetControl('APG_LIBSUP')).PopupMenu := tpopupmenu(GetControl('POP_COM2AFF'));
        TPopUpMenu(GetControl('POP_COM2AFF')).autopopup := true;
        SetControlProperty('APG_LIBSUP','Color',clWindow);
      	SetControlEnabled('APG_LIBSUP',true);
      End
    else
    	Begin
    		SetControlProperty('APG_LIBSUP','Color',clbtnFace);
      	SetControlEnabled('APG_LIBSUP',false);
      End;

// Gestion des seuils en cumul frais ou fourniture
   if ((C.element = 'S20') or (C.element = 'V20')) then
     Begin
       SetControlEnabled('APG_SEUILMINI',true);
       SetControlProperty('APG_SEUILMINI','Color',clWindow);
       SetControlEnabled('APG_SEUILMAXI',true);
       SetControlProperty('APG_SEUILMAXI','Color',clWindow);
       SetControlEnabled('TAPG_SEUILMINI',true);
       SetControlEnabled('TAPG_SEUILMAXI',true);
     End
   else
     Begin
       SetControlEnabled('APG_SEUILMINI',false);
       SetControlProperty('APG_SEUILMINI','Color',clbtnFace);
       SetControlEnabled('APG_SEUILMAXI',false);
       SetControlProperty('APG_SEUILMAXI','Color',clbtnFace);
       SetControlEnabled('TAPG_SEUILMINI',false);
       SetControlEnabled('TAPG_SEUILMAXI',false);
     End
   End;

   // Gestion editions detail Pres/frai/four en cumul
   if (zelt2 = '2') or (zelt2 = '3') then
    Begin
      SetControlEnabled('APG_DETAIL',true);
      if (C.detail = 'X') then
      Begin
        SetControlChecked ('APG_DETAIL',true);
      End
      else
      Begin
        SetControlChecked ('APG_DETAIL',false);
      End
    End
    else
    Begin
      SetControlEnabled ('APG_DETAIL',false);
    End;


END;

procedure TOF_AFPROFILGENER.Init_Liste(Sender : TObject);
Begin
  tlistecol(sender).article:='';
  tlistecol(sender).codearticle:=' ';
  tlistecol(sender).libart:='';
  tlistecol(sender).libsup:='';
  tlistecol(sender).seuilmini:=0;
  tlistecol(sender).seuilmaxi:=0;
  tlistecol(sender).cumul:=false;
  tlistecol(sender).replibech:='-';
  tlistecol(sender).detail:='-';
End;

procedure TOF_AFPROFILGENER.RecupLibAcpte(valeur ,zone : string);
Var Formule : THEdit;
BEGIN
  Formule := THEdit(GetControl (zone));
  if Formule = Nil then Exit;
  Formule.SelText:='['+valeur+']';
END;

function TOF_AFPROFILGENER.ctrlValidite : boolean;
Var C,C1 : TListeCol ;
    wi : integer;
    Celt , celt1 ,celt2 : string;     // élément à integrer
    C1elt , c1elt1 ,c1elt2 : string;  // éléments dèjà sélectionnés

BEGIN
  result := true;
  C:=TListeCol(FDispo.Items.Objects[FDispo.ItemIndex]) ;
  Celt  :=  C.element;
  Celt1 :=  C.element[1];
  Celt2 :=  C.element[2];

     // on ne peut pas avoir  un regroupement ou stot par sous-affaire
     // ou reprise de toutes les lignes à l'identique
    //  si on est en présentation cumulée des affaire
    // ce test est en dehaors de la boucle, pour gérer le traiter même si la
    // liste est vide   (FAffiche.Items.Count=0)
    if ((Celt = 'AST') or (Celt = 'ACO') or (Celt = 'TOU') or (Celt = 'XPO'))  then
    Begin
    if  (TCheckBox(GetControl('APG_CUMUL')).checked) Then
      Begin
      Result := false;
      End;
    End;

    For wi:=0 to FAffiche.Items.Count-1 do
      begin
      C1:=TListeCol(FAffiche.Items.Objects[wi]) ;
      C1elt  :=  C1.element;
      C1elt1 :=  C1.element[1];
      C1elt2 :=  C1.element[2];
      // Si Pxx on ne peut pas avoir une autre Pxx (Si on détail , on ne peut pas avoir cumul)
      if ((Celt1 = 'P') or (Celt1 = 'S') or (Celt1 = 'V') or (Celt1 = 'N')) then
        Begin
        if (Celt1 = C1elt1) then
          Begin
          Result := false;
          break;
          End;
        End;

        // on ne peut pas avoir un détail ou cumul des prestations
        //  si on a  detail des employés par prestation
        if (Celt1 = 'P') then
        Begin
        if (C1elt = 'EPR')  then
          Begin
          Result := false;
          break;
          End;
        End;

        // on ne peut pas avoir detail des employés par prestation
        //  si on a  un détail ou cumul des prestations OU un regroupement employé
        if (Celt = 'EPR')  then
        Begin
        if ((C1elt1 = 'P') or (C1elt1 = 'E')) then
          Begin
          Result := false;
          break;
          End;
        End;
        // on ne peut pas avoir  un regroupement ou stot employé
        //  si on a  detail des employés par prestation
        if ((Celt = 'EST') or (Celt = 'ECO'))  then
        Begin
        if (C1elt = 'EPR')  then
          Begin
          Result := false;
          break;
          End;
        End;
         // on ne peut pas avoir la reprise intégrale des lignes
        //  si on a  autre chose que commentaire affaire  ou stot sur client
        if (Celt = 'TOU')  then
        Begin
        if (C1elt <> 'C10') and (C1elt <> 'XCL')  then
          Begin
          Result := false;
          break;
          End;
        End;

        // on ne peut pas avoir   sous total par client
        //  si on a  autre chose la reprise intégrale des lignes
        // a voir plus tard
        if (Celt = 'XCL')  then
        Begin
        if (C1elt <> 'TOU') and (C1elt <> 'C10')  then
          Begin
          Result := false;
          break;
          End;
        End;
        // on ne peut pas autre chose que commentaire affaire
        //  si on a  la reprise intégrale des lignes
        if (Celt <> 'C10') and  (Celt <> 'XCL') then
        Begin
        if (C1elt = 'TOU')  then
          Begin
          Result := false;
          break;
          End;
        End;

        // on ne peut pas avoir de sous-total EMployé ou Prestation
        // Si on a un cumul prest/frais ou four
        if ((Celt = 'EST') or (Celt = 'AST')) then
        Begin
        if ( (C1elt1 <> 'C') and
             (C1elt2 = '1') or (C1elt2 = '2') or (C1elt2 = '3')or (C1elt = 'EST') or (C1elt = 'AST')) then
          Begin
          Result := false;
          break;
          End;
        End;
        // on ne peut pas avoir de cumul prest/frais ou four
        // Si on a sous-total EMployé ou Prestation
        if ( (Celt1 <> 'C') and
             (Celt2 = '1') or (Celt2 = '2') or (Celt2 = '3')) then
        Begin
        if ((C1elt = 'EST') or (C1elt = 'AST')) then
          Begin
          Result := false;
          break;
          End;
        End;

      End; //for
      if (not result) then  PGIInfoAf(TexteMsgAffaire[1],titre);

END;

/////////////////////////////////////////////
// ****** TOT AFPROFILGENER *****************
/////////////////////////////////////////////

procedure TOT_AFPROFILGENER.OnDeleteRecord;
var code,lib,stit: string;
    resa,resp : boolean;
BEGIN
  Code := getfield('CC_CODE');
  Lib :=  RechDom('AFPROFILGENER',Code,False);
  stit := format('Profil facture : %s - %s',[Code,Lib]);
    // controle des liens avec  la table Affaire  et Profilgener
  resa := SupTablesLiees ('AFFAIRE', 'AFF_PROFILGENER', Code, '' , false);
  resp := SupTablesLiees ('PROFILGENER', 'APG_PROFILGENER', Code, '' , false);
  if ((resp) or (resa)) then
  Begin
      PGIBoxAF('Suppression interdite, ce code profil est utilisé',stit);
      lasterror:=1;
      exit;
  End;

END;

procedure AGLTrtDetail( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_AfProfilGener) then TOF_AfProfilGener(MaTOF).TrtDetail else exit;
end;

procedure AGLTrtReplibech( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_AfProfilGener) then TOF_AfProfilGener(MaTOF).TrtReplibech else exit;
end;

procedure AGLTrtCumul( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_AfProfilGener) then TOF_AfProfilGener(MaTOF).TrtCumul else exit;
end;


procedure TOT_AFPROFILGENER.OnLoad;
var st : string;
begin
 st := getfield('CC_CODE');
END;

Procedure AFLanceFiche_ProfilGener(Argument:string);
begin
AGLLanceFiche ('AFF','AFPROFILGENER','','',Argument);
end;

Initialization
//=== TOT =================
registerclasses([TOT_AFPROFILGENER]) ;
registerclasses([TOF_AFPROFILGENER]);
RegisterAglProc( 'TrtReplibech',True,0,AGLTrtReplibech);
RegisterAglProc( 'TrtDetail',True,0,AGLTrtDetail);
RegisterAglProc( 'TrtCumul',True,0,AGLTrtCumul);
end.

