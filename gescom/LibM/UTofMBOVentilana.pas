unit UTofMBOVentilana;

interface
uses  StdCtrls,Controls,Classes,Forms,SysUtils,ComCtrls,Graphics,
{$IFDEF EAGLCLIENT}

{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db, 
{$ENDIF}
{$IFDEF AFFAIRE}
      TraducAffaire,
{$ENDIF}
      HEnt1,HCtrls,UTOF,Vierge,UTob,EntGC,Dicobtp,UtilGC,M3FP, ParamSoc,UtilPgi;

Type TListeCol = Class
  champs, Lib, IsLib, Constante : String;
  position , longueur : integer;
  sSection : Boolean;
end ;

     TOF_VENTILANA = Class (TOF)
     public
        Faffiche, FDispo, FafficheLib : TListBox;
        procedure OnUpdate ; override ;
        procedure OnDelete ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnClose;  override ;

     private
        TOBAffich, TOBDispo, TOBAffichLib, TOBDispoLib : TOB;
        CodeEtab, CodeAxe, CodeType : String;
        bModif : Boolean;
        // MAJ Table / ecran
        procedure AskMajVentil;
        procedure MajVentil;
        procedure CreatTobStruc (TObMaj:TOB; NumItem,Cpt:integer; TraiteLib:Boolean);
        procedure LoadVentil;
        procedure ChargeItem (TObDet: TOB; NumItem:integer; TraiteLib,Affich:Boolean);
        function TraiteChampLibre(Libelle:String):String ;
        Procedure GetAxeParDefaut ;
        // Gestion des boutons de déplacement
        procedure MoveItem(Sens : String);
        procedure BUp(TraiteLib : boolean);
        procedure BDown(TraiteLib : boolean);
        procedure BLeft(TraiteLib : boolean);
        procedure BRight(TraiteLib : boolean);
        function CtrlValidite(TraiteLib : boolean) : boolean;
        procedure ClearEcran;
        function MajPosition (TraiteLib : boolean) : boolean;
        function AdapteLngChamps (Lng :integer ; Champs : string ) : integer;
        function AdapteLibelleItems (wCode,wlib : string) : string;
        // Chargement Allocation
        procedure ToutAllouer;
        procedure ToutLiberer;
        procedure ChargeLesListBox ;
        procedure SaisieEnabled ( TLB : TListBox ) ;
        procedure ChargeLesTob (Traitelib,TobAffichOnly : Boolean) ;
        // Evenements
        procedure FAfficheClick(Sender: TObject);
        procedure FAfficheLibClick(Sender: TObject);
        Procedure LockUnlockConstante ( champ, lib : string ; C : TListeCol ) ;
        Procedure FRecupLng(TraiteLib : boolean); // modif longueur
        Procedure FRecupsSection; // modif creat code sous section
        Procedure OnExitConstante (Sender: TObject) ;
        END ;


     const
	// libellés des messages de la TOF  afprofilgener
	TexteMsgVentil: array[1..3] of string 	=
          (
          {1}        'La longueur maximum du champ est dépassée'
          {2}       ,'Le code de la section analytique ne peut être supérieur à 17 car.'
          {3}       ,'Le libellé de la section analytique ne peut être supérieur à 35 car.'
          );

procedure AGLMoveItem( parms: array of variant; nb: integer );
procedure AGLLoadVentil( parms: array of variant; nb: integer );

implementation

{ TOF_AFVENTILANA}


procedure TOF_VENTILANA.OnClose;
Begin
  AskMajVentil;
  Toutliberer;
End;

procedure TOF_VENTILANA.OnArgument(stArgument : String );
Var Combo   : THValComboBox;
BEGIN
  Inherited;
  // Condition plus des axes / nombre d'axes utilisés
  Combo := THValComboBox (GetControl ('GDA_AXE') );
//  if VH_GC.GCVentAxe5 then Combo.Plus := ' AND (X_AXE= "A1" or X_AXE= "A2" or X_AXE= "A3" or X_AXE= "A4" or X_AXE= "A5")' else
//  if VH_GC.GCVentAxe4 then Combo.Plus := ' AND (X_AXE= "A1" or X_AXE= "A2" or X_AXE= "A3" or X_AXE= "A4")' else
  if VH_GC.GCVentAxe3 then Combo.Plus := ' AND (X_AXE= "A1" or X_AXE= "A2" or X_AXE= "A3")' else
  if VH_GC.GCVentAxe2 then Combo.Plus := ' AND (X_AXE= "A1" or X_AXE= "A2")' else
                           Combo.Plus := ' AND X_AXE= "A1"';

  SetFocusControl ('GDA_AXE');
  Combo.Value := 'A1' ;
  CodeAxe := 'A1' ;

  GetAxeParDefaut ;

  with THValComboBox(GetControl('GDA_TYPECOMPTE')) do  //Affaire - Lien Paie
   if GetParamSoc('SO_AFLIENPAIEANA')
    then Plus := 'AND (CO_LIBRE LIKE "%COMPTA%" OR CO_LIBRE LIKE "%PAIE%") AND CO_CODE <> "VST"';

  Faffiche   := TListBox   (GetControl('FAFFICHE'));
  FDispo     := TListBox   (GetControl('FDISPO'));
  FafficheLib:= TListBox   (GetControl('FAFFICHELIB'));
  Faffiche.OnClick := FAfficheClick; FafficheLib.OnClick := FAffichelibClick;

  // sortie du champ constante
  THEdit (GetControl('GDA_CSTTE')).OnExit := OnExitConstante ;
  THEdit (GetControl('GDA_CSTTELIB')).OnExit := OnExitConstante ;

  ToutAllouer;
  ChargeLesTob(False,False);
  ChargeLesTob(True,False);
  ChargeLesListBox;
  SaisieEnabled ( FAffiche ) ;
  SaisieEnabled ( FAfficheLib ) ;
END;

procedure TOF_VENTILANA.OnExitConstante ( Sender : TObject ) ;
var C : TListeCol ;
    T : TListBox ;
    S : string ;
begin
S := THEdit ( Sender ) .Name  ;
if S = 'GDA_CSTTE' then T := FAffiche else T := FAfficheLib ;
if T.ItemIndex = -1 then exit ;
C := TListeCol ( T.Items.Objects [ T.ItemIndex ] );
if pos ( 'constante', LowerCase ( C.champs ) ) <> 0 then
   if  C.Constante <> THEDit ( GetControl ( S ) ).Text then
       begin
       C.Constante := THEDit ( GetControl ( S ) ).Text ;
       bModif := true ;
       end ;
end ;


procedure TOF_VENTILANA.OnUpdate;
Var Q : TQuery ;
BEGIN
  Transactions(MajVentil,2);
  // mcd 19/06/01 car si modif, table non rechargée
  VH_GC.MGCTOBAna.ClearDetail;
  Q:=OpenSQL('SELECT * FROM DECOUPEANA',True,-1,'DECOUPEANA') ;
  if Not Q.EOF then VH_GC.MGCTOBAna.LoadDetailDB('DECOUPEANA','','',Q,False) ;
  Ferme(Q);
  BModif:=False ;
END;

procedure TOF_VENTILANA.AskMajVentil;
BEGIN
  if (bModif) then
     If (PGIAskAF('Voulez vous enregistrer la ventilation analytique','')= mrYes) then
        Transactions(MajVentil,2);
END;

procedure TOF_VENTILANA.MajVentil;
Var Req : string;
    TobMaj : TOB;
    Cpt,wi : Integer;
BEGIN
  // supression des éléments existants (code + libellé)
  Req := 'DELETE FROM DECOUPEANA WHERE GDA_AXE = "'+CodeAxe +
         '" AND GDA_TYPECOMPTE ="' + CodeType +
         '" AND GDA_ETABLISSEMENT ="' + CodeEtab + '"';
  ExecuteSQL(Req);
  TobMaj := TOB.create('Creat Ventil',NIL,-1);
  Cpt := 0;
  // Création des élements du code section
  For wi:=0 to FAffiche.Items.Count-1 do
      BEGIN
      Inc(Cpt);
      CreatTobStruc(TObMaj, wi, Cpt, False);
      END ;
  // Création des élements du libellé section
  Cpt := 0;
  For wi:=0 to FAfficheLib.Items.Count-1 do
      BEGIN
      Inc(Cpt);
      CreatTobStruc(TObMaj, wi, Cpt, True);
      END ;
  TOBMaj.InsertOrUpdateDB(False) ;
  TobMaj.free;
END;

procedure TOF_VENTILANA.CreatTobStruc (TObMaj : TOB; NumItem, Cpt : integer; TraiteLib : Boolean);
Var TobDetMaj : TOB;
    C : TlisteCol;
BEGIN
  if TraiteLib then C:=TListeCol(FAfficheLib.Items.Objects[NumItem])
               else C:=TListeCol(FAffiche.Items.Objects[NumItem]) ;
  TobDetMaj := Tob.create('DECOUPEANA',TobMaj,-1);
  TobDetMaj.PutValue('GDA_AXE',CodeAxe);
  TobDetMaj.PutValue('GDA_TYPECOMPTE',CodeType);
  TobDetMaj.PutValue('GDA_ETABLISSEMENT',CodeEtab);
  if TraiteLib then TobDetMaj.PutValue('GDA_TYPESTRUCTANA','LSE')
               else TobDetMaj.PutValue('GDA_TYPESTRUCTANA','SEC');
  TobDetMaj.PutValue('GDA_RANG',Cpt);
  TobDetMaj.PutValue('GDA_LIBCHAMP',C.Champs);
  TobDetMaj.PutValue('GDA_ISLIBELLE',C.IsLib) ;
  TobDetMaj.PutValue('GDA_POSITION',C.Position);
  TobDetMaj.PutValue('GDA_LONGUEUR',C.Longueur);
  TobDetMaj.PutValue('GDA_CSTTE',C.Constante) ;
  if C.sSection then TobDetMaj.PutValue('GDA_CREATSSSECTION','X')
                else TobDetMaj.PutValue('GDA_CREATSSSECTION','-');
END;

procedure TOF_VENTILANA.LoadVentil;
Var CCAxe  : THValComboBox ;
    CCType : THValComboBox ;
    CCEtab : thvalcombobox ;
BEGIN
// Récupération des composants utilisés
  CCAxe := THValComboBox (GetControl ('GDA_AXE') );
  if (CCAxe.Text = '')  then
     PgiBoxAF('Vous devez saisir un axe analytique',TitreHalley) else
     BEGIN
     AskMajVentil;
     ToutLiberer;  // libération des tobs
     Toutallouer;  // création des tobs
     ClearEcran;
     ChargeLesTob(False,False);
     ChargeLesTob(True,False); //Chargement pour le code + lib section
     ChargeLesListBox;
     SaisieEnabled ( FAffiche ) ;
     SaisieEnabled ( FAfficheLib ) ;
     CodeAxe := CCAxe.value;
     CCType := THValComboBox (GetControl('GDA_TYPECOMPTE'));
     CodeType := CCType.Value ;
     CCEtab := thvalcombobox (GetControl('GDA_ETABLISSEMENT'));
     CodeEtab := CCEtab.Value ;
     MajPosition(True); MajPosition(False);
     bModif := false;
     END ;
END;


procedure TOF_VENTILANA.OnDelete;
Var Req : string ;
BEGIN
// supression des éléments existants
If (PGIAskAF('Confirmer-vous la suppression de votre ventilation analytique ?','Ventilation analytique')= mrYes) then
   BEGIN
   Req := 'DELETE FROM DECOUPEANA WHERE GDA_AXE = "'+CodeAxe +
          '" AND GDA_TYPECOMPTE ="' + CodeType +
          '" AND GDA_ETABLISSEMENT ="' + CodeEtab + '"';
   ExecuteSQL(Req);
   CodeAxe := 'A1' ;
   GetAxeParDefaut ;
   LoadVentil ;
   END;
END;

procedure TOF_VENTILANA.ToutAllouer;
BEGIN
  TOBAffich:=TOB.Create('ventil Affich',Nil,-1) ;
  TOBDispo:=TOB.Create('ventil Dispo',Nil,-1) ;
  TOBAffichLib:=TOB.Create('ventil Affich Lib',Nil,-1) ;
  TOBDispoLib:=TOB.Create('ventil Dispo Lib',Nil,-1) ;
END;

procedure TOF_VENTILANA.Toutliberer;
Var i : integer; 
BEGIN
  if  FDispo <> Nil then
      for i:=0 to FDispo.Items.Count-1 do
      begin
      	TRY

        if (FDispo.Items.Objects[i] <> nil) and (FDispo.Items.Objects[i] IS TListeCol) then
        begin
             TListeCol(FDispo.Items.Objects[i]).Free ;
        	FDispo.Items.Objects[i] := nil ;
        end;
        Except
          ;
        END;
      end;
  if  FAffiche <> Nil then
      for i:=0 to FAffiche.Items.Count -1 do
      begin
      	TRY
        if (FAffiche.Items.Objects[i] <> nil)  and (FAffiche.Items.Objects[i] IS TListeCol)  then
        begin
             TListeCol(FAffiche.Items.Objects[i]).Free ;
        	FAffiche.Items.Objects[i]:= nil ;
        end;
        except
          ;
        END;
      end;
  if  FAfficheLib <> Nil then
      for i:=0 to FAfficheLib.Items.Count-1 do
      begin
        Try
        if (FAfficheLib.Items.Objects[i] <> nil)  and (FAfficheLib.Items.Objects[i] IS TListeCol) then
        begin
             TListeCol(FAfficheLib.Items.Objects[i]).Free ;
        	FAfficheLib.Items.Objects[i]:= nil ;
        end;
        Except
          ;
        end;
      end;
  TOBAffich.Free ; TOBAffich:=Nil ; TOBAffichLib.Free ; TOBAffichLib:=Nil ;
  TOBDispo.Free ; TOBDispo:=Nil ; TOBDispoLib.Free ; TOBDispoLib:=Nil ;
  (*
  Faffiche.Items.Clear ; FDispo.Items.Clear; FafficheLib.Items.Clear;
  Faffiche.Items.Capacity := 0;
  FDispo.Items.Capacity := 0;
  FafficheLib.Items.Capacity := 0;
  *)
END;

procedure TOF_VENTILANA.ClearEcran;
Var CBox : TCheckBox;
BEGIN
  Fdispo.clear ; Faffiche.clear ; FafficheLib.clear ;
  //
  Faffiche.Items.Capacity := 0;
  FDispo.Items.Capacity := 0;
  FafficheLib.Items.Capacity := 0;
  //
  SetControlText('GDA_LONGUEUR','') ; SetControlText('GDA_POSITION','') ;
  SetControlText('LONGUEURTOT','') ; SetControlText('GDA_CSTTE','') ;
  SetControlText('GDA_LONGUEURLIB','') ; SetControlText('GDA_POSITIONLIB','') ;
  SetControlText('LONGUEURTOTLIB','') ; SetControlText('GDA_CSTTELIB','') ;
  CBox := TCheckBox(GetControl('GDA_CREATSSSECTION')) ; CBox.Checked := False ;
END;

procedure TOF_VENTILANA.ChargeLesListBox;
Var  wi : integer;
      TobDet : Tob;
BEGIN
  for wi:=0  to TobDispo.Detail.count-1 do
      BEGIN  // boucle sur les éléments disponibles
      TobDet := TobDispo.Detail[wi];
      ChargeItem (TObDet, wi,  False,False);
      END ;
  for wi:=0  to TobAffich.Detail.count-1 do
      BEGIN  // boucle sur les éléments affichés
      TobDet := TobAffich.Detail[wi];
      ChargeItem (TObDet, wi,  False,True);
      END ;
  if TobAffich.detail.count > 0 then
  begin
  	FAffiche.ItemIndex := 0;
    FAfficheClick(Self);
  end;
  for wi:=0  to TobAffichLib.Detail.count-1 do
      BEGIN  // boucle sur les éléments affichés
      TobDet := TobAffichLib.Detail[wi];
      ChargeItem (TObDet, wi,  True,True);
      END;
  if TobAffichLib.detail.count > 0 then
  begin
  	FafficheLib.ItemIndex := 0;
    FAfficheLibClick(Self);
  end;
END;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 26/06/2002
Modifié le ... : 27/06/2002
Description .. : Activation / Désactivation des contrôles de la fiche
Mots clefs ... : BOS5;INTERCOMPTA;
*****************************************************************}
Procedure TOF_VENTILANA.SaisieEnabled ( TLB : TListBox ) ;
var st   : string ;
BEGIN
st := '' ;
if TLB = FAfficheLib then st := 'LIB' ;
if TLB.Items.Count = 1 then
   begin
   SetControlEnabled ( 'BLEFT' + st, true ) ;
   SetControlEnabled ( 'BUP' + st, false ) ;
   SetControlEnabled ( 'BDOWN' + st, false ) ;
   end ;
if TLB.Items.Count > 1 then
   begin
   SetControlEnabled ( 'BLEFT' + st, true ) ;
   SetControlEnabled ( 'BUP' + st, true ) ;
   SetControlEnabled ( 'BDOWN' + st, true ) ;
   end ;
if TLB.Items.Count = 0 then
   begin
   SetControlEnabled ( 'BLEFT' + st, false ) ;
   SetControlEnabled ( 'BUP' + st, false ) ;
   SetControlEnabled ( 'BDOWN' + st, false ) ;

        SetControlText ('GDA_LONGUEUR' + st, '0') ;
        SetControlText ('GDA_CSTTE' + st, '') ;
        SetControlEnabled ('GDA_CSTTE' + st, False) ;
        SetControlProperty ('GDA_CSTTE' + st, 'Color', clBtnFace ) ;
   end ;
END;

procedure TOF_VENTILANA.ChargeItem (TObDet:TOB; NumItem:integer; TraiteLib,Affich:Boolean);
Var C : TlisteCol;
BEGIN
  C:= TlisteCol.Create;
  C.Position := TobDet.GetValue('GDA_POSITION');
  C.Longueur := TobDet.GetValue('GDA_LONGUEUR');
  C.Champs   := TobDet.GetValue('GDA_LIBCHAMP');
  C.IsLib := TobDet.GetValue('GDA_ISLIBELLE');
  C.Lib := TobDet.GetValue('_LIB');
  C.Constante := TobDet.GetValue('GDA_CSTTE') ;
  if Affich then C.sSection := (TobDet.GetValue ('GDA_CREATSSSECTION') = 'X');
  if TraiteLib then
     BEGIN
     if Affich then FAfficheLib.Items.AddObject(C.lib, C) else
        FDispo.Items.AddObject(C.lib, C);
     END else
     BEGIN
     if Affich then FAffiche.Items.AddObject(C.lib, C)else
        FDispo.Items.AddObject(C.lib, C);
     END ;
END;

procedure TOF_VENTILANA.ChargeLesTob (Traitelib, TobAffichOnly : Boolean) ;
Var QQ : TQuery ;
    Req, wcode, wlib, stWhere, Typestr, TypeTT : string ;    
    nbaff, i, Taille : integer ;
    j : integer;
    Tobdet, TobDetSuiv, TobDetDispo, TobTrt : TOB ;
    CCAxe, CCType : THValComboBox ;
    CCEtab : thvalcombobox ;
    Typ, NomChamp, Lib,From,libChamp: String ;
BEGIN
  TypeTT:='BAA' ;
  TobTrt := Nil;  //mcd 12/02/03
  if TraiteLib then TypeStr:='LSE' else Typestr:='SEC';

  if Not (TobAffichOnly) then
     BEGIN
     CCAxe  := THValComboBox(GetControl('GDA_AXE'));
     CCEtab := thvalcombobox(GetControl('GDA_ETABLISSEMENT'));
     CCType := THValComboBox(GetControl('GDA_TYPECOMPTE'));
     Req := 'SELECT * FROM DECOUPEANA WHERE GDA_TYPESTRUCTANA="'+ TypeStr +
            '" AND GDA_AXE = "'+CCAxe.Value +
            '" AND GDA_ETABLISSEMENT ="' + CCEtab.Value +
            '" AND GDA_TYPECOMPTE ="' + CCType.Value + '"' ;
     QQ:=OpenSQL(Req,True);
     if Not QQ.EOF then
        BEGIN
        if Traitelib then
           BEGIN
           TOBAffichLib.LoadDetailDB('DECOUPEANA','','',QQ,False) ;
           TOBAffichLib.Detail[0].AddChampSup('_LIB',True) ;
           TOBAffichLib.AddChampSup('_ISLIBELLE',True) ;
           TOBAffichLib.AddChampSup('_CONSTANTE',True) ;
           nbaff := TOBAffichLib.Detail.count;
           END else
           BEGIN
           TOBAffich.LoadDetailDB('DECOUPEANA','','',QQ,False) ;
           TOBAffich.Detail[0].AddChampSup('_LIB',True) ;
           TOBAffich.AddChampSup('_ISLIBELLE',True) ;
           TOBAffich.AddChampSup('_CONSTANTE',True) ;
           nbaff := TOBAffich.Detail.count;
           END;
        END else
        nbaff := 0;
     Ferme(QQ) ;
     END else
     nbaff := 0;

  // Recup des donnees de la tablette
     // mcd 21/11/02 StWhere:= 'AND CO_LIBRE LIKE "%MODE%"' ;
  if CCType.Value <> 'CON' then
  begin
  StWhere:= 'AND CO_LIBRE LIKE "%'+VH_GC.GCMarcheVentilAna+'%"' ;
  end else
  begin
  	StWhere:= 'AND CO_LIBRE LIKE "%;CON;%"' ;
  end;
  Req := 'SELECT CO_CODE,CO_LIBELLE,CO_LIBRE FROM COMMUN WHERE CO_TYPE="'+TypeTT+'" ' + stWhere + ' Order by CO_CODE';
  QQ:=OpenSQL(Req,True) ;
  If not QQ.EOF Then
  begin
    TobTrt := Tob.create ('liste champs', Nil,-1);
    TobTrt.LoadDetailDB('','','',QQ,False);
  end ;
  Ferme(QQ) ;

  For i := 0 To TobTrt.detail.count-1 do
     BEGIN
     wLib:='' ;
     wcode:= TobTrt.detail[i].GetValue('CO_CODE');
     TobTrt.detail[i].AddChampSup('_LIBCHAMP',False) ;
     Lib:=TobTrt.detail[i].GetValue('CO_LIBELLE') ;
     ReadTokenSt(Lib) ;
     if ReadTokenSt(Lib)='&#@' then wlib:=TraiteChampLibre(Lib) ;
     NomChamp:=TobTrt.detail[i].GetValue('CO_LIBRE') ;
     TobTrt.detail[i].PutValue('_LIBCHAMP',ReadTokenSt(NomChamp)) ;
     TobTrt.detail[i].AddChampSup('_TYPECHAMP',False) ;
     TobTrt.detail[i].AddChampSup('_ISLIBELLE',False) ;
     TobTrt.Detail[i].PutValue('_ISLIBELLE','-') ;
     QQ:=OpenSQL('SELECT DH_TYPECHAMP, DH_LIBELLE, (SELECT DT_NOMTABLE FROM DETABLES WHERE DT_PREFIXE=DH_PREFIXE) AS NOMTABLE FROM DECHAMPS WHERE DH_NOMCHAMP="'+TobTrt.detail[i].GetValue('_LIBCHAMP')+'"',True) ;
     if not QQ.EOF then
        BEGIN
        TobTrt.detail[i].PutValue('_TYPECHAMP',QQ.FindField('DH_TYPECHAMP').AsString) ;
         //if wlib='' then wlib:=QQ.FindField('DH_LIBELLE').AsString ;
         if wlib='' then begin
            wlib:=TobTrt.detail[i].GetValue('CO_LIBELLE');
            j:=pos(';',wlib);
            if j >0 then wlib:=copy(wlib,1,J-1);
{$IFDEF AFFAIRE}
            wlib:='('+QQ.FindField('NOMTABLE').AsString+') '+TraduitGa (wlib);
{$ENDIF}
            end;
        END ;
     Ferme (QQ) ;
     Typ:=TobTrt.detail[i].GetValue('_TYPECHAMP') ;
     TobTrt.detail[i].AddChampSup('_TAILLE',False) ;
     //
     // Traitement des constantes
     //
     if (wCode = 'CT1') or (wCode = 'CT2') or (wCode = 'CT3') then
        BEGIN
        Typ := 'VARCHAR(35)' ;
        TobTrt.detail[i].PutValue('_TYPECHAMP', Typ) ;
        wLib := TobTrt.detail[i].GetValue('CO_LIBELLE') ;
        END ;
     Taille := 0 ; // OT : 15/04/2002
     if (copy(Typ,1,7)='VARCHAR') then Taille:=StrToInt(copy(Typ,9,length(Typ)-9))
        else if (Typ='DOUBLE') or (Typ='RATE') then Taille:=12
        else if (Typ='INTEGER') or (Typ='SMALLINT') then Taille:=8
        else if Typ='DATE' then Taille:=10
        else if Typ='COMBO' then Taille:=3 ;
     if ReadTokenSt(NomChamp)='X' then
        BEGIN
        wLib := AdapteLibelleItems (wCode,wlib);
        Taille:=35 ;
        TobTrt.Detail[i].PutValue('_ISLIBELLE','X') ;
        END ;
     TobTrt.detail[i].PutValue('_TAILLE',Taille) ;
     if nbaff <> 0 then  // Recherche si l'element est dans la tob Affiche
        BEGIN
        if Traitelib then
           begin
           TobDet := TobAffichlib.FindFirst(['GDA_LIBCHAMP'],[TobTrt.detail[i].GetValue('_LIBCHAMP')],False);
           if (TobDet<>Nil) and (TobTrt.detail[i].GetValue('_ISLIBELLE') <> 'X') then
              begin
              TobDetSuiv := TobAffichlib.FindNext(['GDA_LIBCHAMP'],[TobTrt.detail[i].GetValue('_LIBCHAMP')],False);
              if TobDetSuiv<>Nil then TobDet:=TobDetSuiv
              else TobDet:=Nil;
             end;
           end
        else
           begin
           TobDet := TobAffich.FindFirst(['GDA_LIBCHAMP'],[TobTrt.detail[i].GetValue('_LIBCHAMP')],False);
           if (TobDet<>Nil) and (TobTrt.detail[i].GetValue('_ISLIBELLE') = 'X') then
              begin
              TobDetSuiv := TobAffich.FindNext(['GDA_LIBCHAMP'],[TobTrt.detail[i].GetValue('_LIBCHAMP')],False);
              if TobDetSuiv<>Nil then TobDet:=TobDetSuiv
              else TobDet:=Nil;
             end;
           end
        END
     else
        TobDet := NIL;
     if TobDet<>nil then
        begin
        TobDet.PutValue('_LIB', wlib);
        end;
     if (wLib <> '.-') then  // sinon stat non utilisée
        BEGIN
        if traiteLib then TobDetDispo := Tob.create('DECOUPEANA',TobDispoLib,-1) else
                          TobDetDispo := Tob.create('DECOUPEANA',TobDispo,-1);
        TobDetDispo.PutValue('GDA_LIBCHAMP',  TobTrt.detail[i].GetValue('_LIBCHAMP'));
        TobDetDispo.AddChampSup('_LIB',False) ;
        TobDetDispo.PutValue('_LIB', wlib);
        TobDetDispo.PutValue('GDA_LONGUEUR',TobTrt.detail[i].GetValue('_TAILLE'));
        TobDetDispo.PutValue('GDA_ISLIBELLE',TobTrt.Detail[i].GetValue('_ISLIBELLE'));
        END;
     END;
  TobDispo.detail.sort ('_LIB'); //mcd 27/05/03
  TobTrt.Free;
  bModif := False;
END;

 // AC: Retourne le libelle du champ de type combo
function TOF_VENTILANA.TraiteChampLibre(Libelle:String):String ;
var Tablette,Code,St,Prefixe : String ;
BEGIN
  Result:='' ;
  Tablette:=ReadTokenSt(Libelle) ;
  Code:=ReadTokenSt(Libelle) ;
  Prefixe := '';
  if pos('AT',Code) > 0 then Prefixe:= '(ARTICLE) ' else
  if Pos('MT',Code) > 0 then prefixe := '(AFFAIRE) ' else
  if Pos('MR',Code) > 0 then prefixe := '(RESSOURCE) ' else
  if Pos('LF',Code) > 0 then prefixe := '(PARAM) ' else
  if Pos('PS',Code) > 0 then prefixe := '(LIBRE ARTICLE) ' else
  if Pos('FT',Code) > 0 then prefixe := '(LIBRE FOUR) ' else
  if Pos('CR',Code) > 0 then prefixe := '(LIBRE RESSOURCE) ' else
  if Pos('CT',Code) > 0 then prefixe := '(LIBRE CLIENT) ';

//  st := RechDom (Tablette,Code,False);
	if Tablette = 'GCZONELIBRE' then st := Prefixe + RechDomZoneLibre (Code, False)
  														else st := prefixe + RechDom (Tablette,Code,False);

  Result:=st ;
END ;

procedure TOF_VENTILANA.FAfficheClick(Sender: TObject);
Var C : TListeCol ;
    bModifsav : Boolean;
BEGIN
  MajPosition(False);
  if FAffiche.ItemIndex >= 0 then
     BEGIN
     bModifSav := bModif;
     C:=TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]);

     // blocage ou pas de la saisie de la constante
     LockUnlockConstante ( 'GDA_CSTTE', 'TGDA_CSTTE', C ) ;

     SetControlText('GDA_POSITION',IntToStr(C.Position));
     SetControlText('GDA_LONGUEUR',IntToStr (C.Longueur));
     SetControlProperty('GDA_CREATSSSECTION','Checked',C.sSection);
     bModif := bModifSav;
     END ;
END;

procedure TOF_VENTILANA.FAffichelibClick(Sender: TObject);
Var C : TListeCol ;
    bModifsav : Boolean;
BEGIN
  MajPosition(True);
  if FAfficheLib.ItemIndex >= 0 then
     BEGIN
     bModifSav := bModif;
     C:=TListeCol(FAfficheLib.Items.Objects[FAfficheLib.ItemIndex]);

     // blocage ou pas de la saisie de la constante
     LockUnlockConstante  ( 'GDA_CSTTELIB', 'TGDA_CSTTELIB', C ) ;

     SetControlText('GDA_POSITIONLIB',IntToStr(C.Position));
     SetControlText('GDA_LONGUEURLIB',IntToStr (C.Longueur));
     bModif := bModifSav;
     END ;
END;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 25/06/2002
Modifié le ... : 25/06/2002
Description .. : Activation / Désactivation de la constante
Mots clefs ... : BOS5;INTERCOMPTA;
*****************************************************************}
Procedure TOF_VENTILANA.LockUnlockConstante ( champ,lib : string ; C : TListeCol ) ;
begin
if Pos ( 'constante', LowerCase( C.champs ) ) <> 0 then
   begin  // on débloque
   SetControlEnabled ( champ, true ) ;
   SetControlEnabled ( lib, true ) ;
   SetControlProperty ( champ, 'Color', clWindow ) ;
   SetControlText ( champ, C.Constante ) ;
   end else
   begin // on bloque
   SetControlText ( champ, C.Constante ) ;
   SetControlEnabled ( champ, false ) ;
   SetControlEnabled ( lib, false ) ;
   SetControlProperty ( champ, 'Color', clBtnFace ) ;
   end ;
end ;

Procedure TOF_VENTILANA.FRecupsSection ;
Var C : TListeCol ;
    CBox : TCheckBox;
BEGIN
  if FAffiche.ItemIndex < 0 then exit;
  C:=TListeCol(FAfficheLib.Items.Objects[FAfficheLib.ItemIndex]);
  if (C <> nil) then
      BEGIN
      CBox := TCheckBox(GetControl('GDA_CREATSSSECTION'));
      C.sSection := CBox.Checked;
      END;
END;

Procedure TOF_VENTILANA.FRecupLng(TraiteLib : boolean);
Var C : TListeCol ;
    suff : string;
BEGIN
  if Traitelib then
      BEGIN
      if FAfficheLib.ItemIndex < 0 then exit;
      C:=TListeCol(FAfficheLib.Items.Objects[FAfficheLib.ItemIndex]);
      suff :='LIB';
      END else
      BEGIN
      if FAffiche.ItemIndex < 0 then exit;
      C:=TListeCol(FAffiche.Items.Objects[FAffiche.ItemIndex]);
      Suff :='';
      END;
  if C <> nil then
      BEGIN
      C.Longueur := StrToInt(GetControlText('GDA_LONGUEUR'+suff));
      if Not(MajPosition(TraiteLib)) then SetControlText('GDA_LONGUEUR'+Suff,IntToStr(C.Longueur));
      SetControlText('GDA_POSITION'+Suff,IntToStr(C.Position));
      END;
END;

// Recalcul de la position de l'item affiché et contrôles
function TOF_VENTILANA.MajPosition (TraiteLib : boolean) : Boolean;
Var C : TListeCol ;
    i, Posit, numError, max, MaxLng,NumI,Lng : integer ;
    Suff : string;    // NomTTLibre,
    FTraite : TListBox;
    TobComp : Tob ;
    IsConstante : boolean;
BEGIN
  isConstante := false;
  if Traitelib then
      BEGIN
      FTraite := FAfficheLib;
      Suff := 'LIB'; // NomTTLibre := 'GCCHAMPSANA';
      MaxLng := 35;
      END else
      BEGIN
      FTraite := FAffiche;
      Suff := ''; // NomTTLibre := 'GCCHAMPSANA';
      MaxLng := 17;
      END;
  Posit := 1;  numError := 0; Result := true; Max := 0 ;
  For i := 0 to FTraite.Items.Count-1 do
      BEGIN
      C:=TListeCol(FTraite.Items.Objects[i]);
      C.position :=  Posit;
      if copy(C.champs,1,9)='Constante' then IsConstante := true;
      {if TraiteLib then TobP:=TOBAffichLib.FindFirst(['GDA_LIBCHAMP','GDA_ISLIBELLE'],[C.Champs,C.IsLib],False)
                   else TobP:=TOBAffich.FindFirst(['GDA_LIBCHAMP','GDA_ISLIBELLE'],[C.Champs,C.IsLib],False) ;
      if tobP=nil then TobP:=TOBDispo.FindFirst(['GDA_LIBCHAMP','GDA_ISLIBELLE'],[C.Champs,C.IsLib],False) ;  }
      TobComp := TOBDispo.FindFirst(['GDA_LIBCHAMP','GDA_ISLIBELLE'],[C.Champs,C.IsLib],False) ;
      if TobComp <> nil then Max := AdapteLngChamps(TobComp.GetValue('GDA_LONGUEUR'),C.Champs) ;
      if (C.Longueur > Max) then
         BEGIN
         numError := 1;
         C.Longueur := Max ;
         END;
      Posit := C.Position + C.Longueur;
      END;
  Dec(Posit);
  if Posit > MaxLng then
     if TraiteLib then numError := 3 else numError := 2 ;
  if  numError > 0 then
      BEGIN
      Result := False;
      if not IsConstante then PGIBoxAF (TexteMsgVentil[numError],'Ventilation Analytique');
      END;
  if (NumError >= 2) then
      BEGIN
      if FTraite.ItemIndex>=0 then
         BEGIN
         NumI:=FTraite.ItemIndex ;
         C:=TListeCol(FTraite.Items.Objects[NumI]) ;
         Lng := C.Longueur -(Posit - MaxLng);
         if Lng >= 0 then C.Longueur :=Lng;
         Posit := MaxLng;
         END;
      END;
  SetControltext('LONGUEURTOT'+suff ,IntToStr(Posit));
END;

function TOF_VENTILANA.AdapteLngChamps (Lng :integer; Champs : string) : integer;
BEGIN
  Result := Lng;
  if Champs ='AC1' then Result :=  VH_GC.CleAffaire.Co1Lng else
     if Champs ='AC2' then Result :=  VH_GC.CleAffaire.Co2Lng else
        if Champs ='AC3' then Result :=  VH_GC.CleAffaire.Co3Lng else
           Exit;
END;

function TOF_VENTILANA.AdapteLibelleItems (wCode,wlib : string) : string;
Var LibLabel,Codestat,From : string;
BEGIN
	From := '';
  Result := wLib; LibLabel := '';
  // maj direct du libellé
  if wCode ='AC1' then Result := '(AFFAIRE)'+VH_GC.CleAffaire.Co1Lib else
  if wCode ='AC2' then Result := '(AFFAIRE)'+VH_GC.CleAffaire.Co2Lib else
  if wCode ='AC3' then Result := '(AFFAIRE)'+VH_GC.CleAffaire.Co3Lib else
   // Récup du libellé de la stat associée
   if wCode = 'AS1' then begin LibLabel := 'TAFF_LIBREAFF1'; From := 'AFFAIRE'; end else    // Stat Affaire
   if wCode = 'AS2' then begin LibLabel := 'TAFF_LIBREAFF2'; From := 'AFFAIRE'; end else
   if wCode = 'AS3' then begin LibLabel := 'TAFF_LIBREAFF3'; From := 'AFFAIRE'; end else
   if wCode = 'AS4' then begin LibLabel := 'TAFF_LIBREAFF4'; From := 'AFFAIRE'; end else
   if wCode = 'AS5' then begin LibLabel := 'TAFF_LIBREAFF5'; From := 'AFFAIRE'; end else
   if wCode = 'AS6' then begin LibLabel := 'TAFF_LIBREAFF6'; From := 'AFFAIRE'; end else
   if wCode = 'AS7' then begin LibLabel := 'TAFF_LIBREAFF7'; From := 'AFFAIRE'; end else
   if wCode = 'AS8' then begin LibLabel := 'TAFF_LIBREAFF8'; From := 'AFFAIRE'; end else
   if wCode = 'AS9' then begin LibLabel := 'TAFF_LIBREAFF9'; From := 'AFFAIRE'; end else
   if wCode = 'ASA' then begin LibLabel := 'TAFF_LIBREAFFA'; From := 'AFFAIRE'; end else

   if wCode = 'PS1' then begin LibLabel := 'TGA_LIBREART1'; From := 'ARTICLE'; end else // Stat Article
   if wCode = 'PS2' then begin LibLabel := 'TGA_LIBREART2'; From := 'ARTICLE'; end else
   if wCode = 'PS3' then begin LibLabel := 'TGA_LIBREART3'; From := 'ARTICLE';end else
   if wCode = 'PS4' then begin LibLabel := 'TGA_LIBREART4'; From := 'ARTICLE';end else
   if wCode = 'PS5' then begin LibLabel := 'TGA_LIBREART5'; From := 'ARTICLE';end else
   if wCode = 'PS6' then begin LibLabel := 'TGA_LIBREART6'; From := 'ARTICLE';end else
   if wCode = 'PS7' then begin LibLabel := 'TGA_LIBREART7'; From := 'ARTICLE';end else
   if wCode = 'PS8' then begin LibLabel := 'TGA_LIBREART8'; From := 'ARTICLE';end else
   if wCode = 'PS9' then begin LibLabel := 'TGA_LIBREART9'; From := 'ARTICLE';end else
   if wCode = 'PSA' then begin LibLabel := 'TGA_LIBREARTA'; From := 'ARTICLE';end else

   if wCode = 'TS1' then begin LibLabel := 'TYTC_TABLELIBRETIERS1'; From := 'TIERS'; end else // Stat clients
   if wCode = 'TS2' then begin LibLabel := 'TYTC_TABLELIBRETIERS2'; From := 'TIERS'; end else
   if wCode = 'TS3' then begin LibLabel := 'TYTC_TABLELIBRETIERS3'; From := 'TIERS'; end else
   if wCode = 'TS4' then begin LibLabel := 'TYTC_TABLELIBRETIERS4'; From := 'TIERS'; end else
   if wCode = 'TS5' then begin LibLabel := 'TYTC_TABLELIBRETIERS5'; From := 'TIERS'; end else
   if wCode = 'TS6' then begin LibLabel := 'TYTC_TABLELIBRETIERS6'; From := 'TIERS'; end else
   if wCode = 'TS7' then begin LibLabel := 'TYTC_TABLELIBRETIERS7'; From := 'TIERS'; end else
   if wCode = 'TS8' then begin LibLabel := 'TYTC_TABLELIBRETIERS8'; From := 'TIERS'; end else
   if wCode = 'TS9' then begin LibLabel := 'TYTC_TABLELIBRETIERS9'; From := 'TIERS'; end else
   if wCode = 'TSA' then begin LibLabel := 'TYTC_TABLELIBRETIERSA'; From := 'TIERS'; end else

   if wCode = 'F1L' then begin Result :=RechDom('GCLIBFAMILLE','LF1',false); From := 'ARTICLE'; end else
   if wCode = 'F2L' then begin Result :=RechDom('GCLIBFAMILLE','LF2',false); From := 'ARTICLE';end else
   if wCode = 'F3L' then begin Result :=RechDom('GCLIBFAMILLE','LF3',false); From := 'ARTICLE';end else

   if wCode = 'CT1' then begin Result := 'Constante 1'; end else
   if wCode = 'CT2' then begin Result := 'Constante 2'; end else
   if wCode = 'CT3' then begin Result := 'Constante 3'; end else
      Exit;

  if LibLabel <> '' then // lecture du libellé de la stat ...
     BEGIN
     CodeStat := CodeTableLibre_2(LibLabel);
     If (CodeStat <> '') then
     begin
		   Result := RechDomZoneLibre (CodeStat, False);
     end;
     END;

if From <> '' then Result := '('+From+') Lib. '+result
							else Result :='Lib. '+Result;
END;

// Gestion des déplacements
function TOF_VENTILANA.ctrlValidite (TraiteLib : boolean)  : boolean;
Var C,C1 : TListeCol ;
    wi : integer;
    FTraiteDispo,FTraiteAffiche : TListBox;
BEGIN
  if TraiteLib then BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAfficheLib; END
               else BEGIN  FTraiteDispo := FDispo; FTraiteAffiche := FAffiche; END;
  result := true;
  C:=TListeCol(FTraiteDispo.Items.Objects[FTraiteDispo.ItemIndex]) ;
  For wi:=0 to FTraiteAffiche.Items.Count-1 do
      BEGIN
        C1:=TListeCol(FTraiteAffiche.Items.Objects[wi]) ;
        if (C.Champs = C1.Champs) then
        BEGIN
          Result := false;
          exit;
        END ;
      END ;
END;

procedure TOF_VENTILANA.BRight(TraiteLib : boolean) ;
var oldi : integer ;
    C : TListeCol ;
    ok : boolean;
    FTraiteDispo,FTraiteAffiche : TListBox;
    TobDisplay : TOB ;
  BEGIN
  if TraiteLib then BEGIN  FTraiteDispo := FDispo ; FTraiteAffiche := FAfficheLib ; TobDisplay := TobAffichLib END
               else BEGIN  FTraiteDispo := FDispo ; FTraiteAffiche := FAffiche ; TobDisplay := TobAffich ; END;

  if FTraiteDispo.ItemIndex>=0 then
     BEGIN
     OldI:=FTraiteDispo.ItemIndex ;
     C:=TListeCol(FTraiteDispo.Items.Objects[OldI]) ;
     ok := CtrlValidite(TraiteLib);
     if (ok) then
        BEGIN
        FTraiteAffiche.Items.AddObject(FTraiteDispo.Items[OldI],C) ;
        CreatTobStruc(TobDisplay, FTraiteAffiche.Items.Count-1, FTraiteAffiche.Items.Count-1, TraiteLib) ;

        if OldI<FTraitedispo.Items.Count then FTraiteDispo.itemIndex:=OldI else FTraiteDispo.itemIndex:=OldI-1 ;
        FTraiteAffiche.ItemIndex:=FTraiteAffiche.Items.Count-1 ;

        // recopie de la constante
        TListeCol(FTraiteAffiche.Items.Objects[FTraiteAffiche.ItemIndex]).Constante := TListeCol(FTraiteDispo.Items.Objects[FTraiteDispo.itemIndex]).Constante ;

        if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
        END ;
     END ;
  SaisieEnabled ( FTraiteAffiche ) ;
end;

procedure TOF_VENTILANA.BLeft(TraiteLib : boolean) ;
var oldi, i : integer ;
    FTraiteAffiche : TListBox ;
    TLC1, TLC2 : TListeCol ;
    Lib : string ;
    TobDisplay : TOB ;
BEGIN
  Lib := '' ;
  if TraiteLib then BEGIN FTraiteAffiche := FAfficheLib; Lib := 'LIB' ; TobDisplay := TobAffichLib ; END
               else BEGIN FTraiteAffiche := FAffiche ; TobDisplay := TobAffich ;END;

  if FTraiteAffiche.ItemIndex>=0 then
     BEGIN
     OldI:=FTraiteAffiche.ItemIndex ;

     // stockage de la constante
     For i:=0 to FDispo.Items.Count-1 do
         BEGIN
         TLC1 := TListeCol(FDispo.Items.Objects[i]) ;
         TLC2 := TListeCol(FTraiteAffiche.Items.Objects[Oldi]) ;
          if (TLC1.Champs = TLC2.Champs) then
              BEGIN
              TLC1.Constante := TLC2.Constante ;
              break ;
              END ;
          END ;

     FTraiteAffiche.Items.Delete(OldI) ;
     TobDisplay.Detail[OldI].Free ;
     if OldI<FTraiteAffiche.Items.Count then FTraiteAffiche.itemIndex:=OldI else FTraiteAffiche.itemIndex:=OldI-1 ;
     if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
     END ;
  SaisieEnabled ( FTraiteAffiche ) ;
END;

procedure TOF_VENTILANA.BUp(TraiteLib : boolean) ;
var oldi : integer ;
    FTraiteAffiche : TListBox;
BEGIN
  if TraiteLib then BEGIN FTraiteAffiche := FAfficheLib; END  // FTraiteDispo := FDispo;
               else BEGIN FTraiteAffiche := FAffiche; END;    // FTraiteDispo := FDispo;

  OldI:=FTraiteAffiche.ItemIndex ;
  if OldI>0 then
     BEGIN
     if (FTraiteAffiche.Items[OldI-1]='IE_CHRONO') then Exit ;
     FTraiteAffiche.Items.Exchange(OldI,OldI-1) ;
     FTraiteAffiche.ItemIndex:=OldI-1 ;
     END ;
  if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
END ;

procedure TOF_VENTILANA.BDown(TraiteLib : boolean) ;
var oldi : integer ;
    FTraiteAffiche : TListBox;
BEGIN
  if TraiteLib then BEGIN FTraiteAffiche := FAfficheLib; END  // FTraiteDispo := FDispo;
               else BEGIN FTraiteAffiche := FAffiche; END;    // FTraiteDispo := FDispo;

  //if (FAffiche.Items[FAffiche.ItemIndex]='IE_CHRONO') then Exit ;
  OldI:=FTraiteAffiche.ItemIndex ;
  if ((OldI>=0) and (oldI<FTraiteAffiche.items.Count-1)) then
     BEGIN
     FTraiteAffiche.Items.Exchange(OldI+1,OldI) ;
     FTraiteAffiche.ItemIndex:=OldI+1 ;
     END ;
  if TraiteLib then FAfficheLibClick(nil) else FAfficheClick(nil) ;
end;

procedure TOF_VENTILANA.MoveItem(Sens : String);
begin
//if Not(SaisieEncours) then Exit;
  bModif := True;
  Uppercase (Sens);
  if Sens = 'UP'   then BUp(False) else
   if Sens = 'DOWN' then BDown(False) else
    if Sens = 'LEFT' then BLeft(False) else
     if Sens = 'RIGHT' then BRight(False) else
      if Sens ='LNG' then FRecupLng(False) else
       if Sens = 'SSECTION' then FRecupsSection else
        if Sens = 'UPLIB'   then BUp(True) else
         if Sens = 'DOWNLIB' then BDown(True) else
          if Sens = 'LEFTLIB' then BLeft(True) else
           if Sens = 'RIGHTLIB' then BRight(True) else
            if Sens = 'LNGLIB' then FRecupLng(True) ;
END;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 13/05/2002
Modifié le ... : 13/05/2002
Description .. : Chargement d'un axe par défaut avec le code 
Mots clefs ... : BOS5;INTERCOMPTA;
*****************************************************************}
procedure TOF_VENTILANA.GetAxeParDefaut ;
var Q : TQuery ;
begin
  Q := OpenSQL('SELECT GDA_TYPECOMPTE,GDA_ETABLISSEMENT FROM DECOUPEANA WHERE GDA_AXE="' + CodeAxe + '"', False) ;
  if not Q.Eof then
     BEGIN
     CodeType := Q.FindField('GDA_TYPECOMPTE').AsString ;
     CodeEtab := Q.FindField('GDA_ETABLISSEMENT').AsString ;
     END else
     BEGIN
     CodeType := 'VEN' ;
     CodeEtab := '' ;
     END ;
  Ferme (Q) ;
  THValcombobox (GetControl('GDA_ETABLISSEMENT')).Value := CodeEtab ;
  THValComboBox (GetControl('GDA_TYPECOMPTE')).Value := CodeType ;
end ;

procedure AGLMoveItem( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
BEGIN
  F:=TForm(Longint(Parms[0]));
  LaTof := Nil ;
  if (F is TFVierge) then Latof:=TFVierge(F).Latof;
  if (Latof is TOF_VENTILANA) then
      TOF_VENTILANA(LaTof).MoveItem(Parms[1]);
END ;

procedure AGLLoadVentil( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
BEGIN
  F:=TForm(Longint(Parms[0]));
  LaTof := Nil ;
  if (F is TFVierge) then Latof:=TFVierge(F).Latof;
  if (Latof is TOF_VENTILANA) then
      TOF_VENTILANA(LaTof).LoadVentil;
END ;

Initialization
registerclasses([TOF_VENTILANA]);
RegisterAglProc( 'MoveItem',True,1,AGLMoveItem);
RegisterAglProc( 'LoadVentil',True,0,AGLLoadVentil);
end.
