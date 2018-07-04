unit LP_impri;

interface

uses Classes, LP_Print, MC_Comm, LP_Base, MC_Erreur, Controls ;

Type
  TLPPrinter = Class ( TComponent)
     Private
     FPrinterInterne : TLPPrinterInterne ;
     protected
      Function  GetPort         : TComPortNumber ;
      Procedure SetPort         (AValue : TComPortNumber) ;
      Function  GetSqce         (Indice : TLP_SqcesPossibles ; Activer : Boolean ) : String ;
      Function  GetLastLigne    : Integer  ;
      Procedure SetLastLigne    (Value : Integer ) ;
      Function  GetLastError    : TImprimErreurs ;
      Function  GetLargeur      : Integer ;
      Procedure SetLargeur      ( Value : Integer ) ;
      Function  GetInitialise   : Boolean ;
      Procedure SetInitialise   ( Value : Boolean ) ;
      Function  GetSqceInUse    (Indice : TLP_SqcesPossibles) : Boolean ;
      Procedure SetSqceInUse    (Indice : TLP_SqcesPossibles ; Value : Boolean ) ;
      Procedure SetInTestMateriel ( Value : Boolean ) ;
      Function  GetInTestMateriel : Boolean ;
     Public
      Constructor Create (AOwner : TComponent ; AImprimante : String ) ; reintroduce; overload ; Virtual ;
      Destructor  destroy ; override ;
      Procedure   InitLastligne ;
      Procedure   InitNewLigne ;
      Function    InitImprimante ( Var Err : TMC_Err ) : Boolean ;
      Function    ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
      Function    OuvrirImprimante( Var Err : TMC_Err ) : Boolean ;
      Function    FermeImprimante ( Var Err : TMC_Err ) : Boolean ;
      Procedure   FormatPolice ( zn : TLP_ZoneImp) ; //Seq : TLP_SetSqcesPossibles ; Colonne, Longeur : Integer ; Extra : String ) ;
      Function    EcrireLigne(St : String ; Ligne : Integer ) : Boolean ;
      Function    isConnected : Boolean ;
      Function    SautPage : Boolean ;
      Function    DoTranslate ( St : String ) : String ; //XMG 30/03/04
      Procedure   Demarre ;

      Property Port : TComPortNumber  Read GetPort write SetPort default PnNone ;
      Property Sqce    [Indice : TLP_SqcesPossibles ; Activer : Boolean ]  : String  read GetSqce ;

      Property LastLigne      : Integer Read GetLastLigne write SetLastLigne ;
      Property LastError      : TImprimErreurs read GetLastError ;
      property Largeur        : Integer read GetLargeur write SetLargeur ;
      Property Initialise     : Boolean read GetInitialise write Setinitialise ;
      Property InTestMateriel : Boolean            read GetInTestMateriel write SetInTestMateriel ;
     End ;

  PDecla = ^TDecla ;
  TDecla = record
           Champ,
           Libelle,
           TypeCh : String ;
           End ;
  TProcAsigneEventsLP = Procedure ( Ctrl : TControl) of Object ;

{$IFDEF eAGLClient}
{$ELSE}
Function EnregLP ( LPParam : TLPParam ; Base : TLPBase ) : Boolean ;
{$ENDIF}
Function ChargeModeleLP (Var LPParam : TLPParam ; Creer : Boolean ; Base : TLPBase ; pAsigneEventsChamp,pAsigneEventsBande : TProcAsigneEventsLP ; Var Err : TMC_Err ; Var ChangementVersion : Boolean ) : Boolean ;

Function LP_ImporteModeleInterne ( NomFichier : String ; Var LPParam : TLPParam ; Base : TLPBase ; pAsigneEventsChamp,pAsigneEventsBande : TProcAsigneEventsLP ; Var Err : TMC_Err ; Var ChangementVersion : Boolean ) : Boolean ;
Function LP_ExporteModeleInterne ( NomFichier : String ; LPParam : TLPParam ; Base : TLPBase ; Var Err : TMC_Err ) : Boolean ;

Function LP_ExporteModele ( NomFichier : String ; TpEtat,NatureEtat,Code,Lng : String ; Var Err : TMC_Err ) : Boolean ;
{$IFDEF eAGLClient}                    //YCP 19/05/03
{$ELSE}
Function LP_ImporteModele ( NomFichier : String ; Var Err : TMC_Err ; Var ChangementVersion : Boolean ) : Boolean ;
{$ENDIF eAGLClient}
Function ExisteModele(TypeEtat,NatEtat,ModeleEtat : String ; Var Langue :String ; OkLangueDef : Boolean  ) : Boolean ;


implementation

uses Hent1, HCtrls, MC_lib, math, Sysutils, Forms,
{$IFDEF eAGLClient}
    UTob
{$ELSE}
   dbTables, Db
{$ENDIF}
 ;
//////////////////////////////////////////////////////////////////////////////////
Procedure ChercheVersion (Mem : TStream ; VBase,StVersion : String ; var OkVersion : Boolean ; var VImp : Double ) ;
var St       : String ;
    pp,posit : Integer ;
Begin
VImp:=-1 ;
PosIt:=Mem.Position ;
setlength(St,length(StVersion)) ;
mem.Read(St[1],length(StVersion)) ;
pp:=pos(VBase,StVersion) ;
OkVersion:=(copy(St,1,pp-1)=Copy(StVersion,1,pp-1)) ;
if (OkVersion) then VImp:=Valeur(Copy(St,pp,Length(VBase)))
               else Mem.Seek(PosIt,soFromBeginning) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function ConstruitCle (LPParam : TLPParam ; IsEnreg,IsSQL : Boolean ) : String ;
Begin
Result:=Format_String(LPParam.TypeEtat,1)+Format_String(LPParam.NatEtat,3)+Format_string(LPParam.CodeEtat,3)+Copy('PG1SQL',1+3*ord(IsSQL),3) ;
if isEnreg then Result:=Result+Format_String(LPParam.Langue,3) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Const VersionModele = '1.00' ; //<1.00.- Aucun controle de version.... ; 1.00 instauration control version Modèle ;
      ConstVModele  = 'Track modèle impresion texte Version '+VersionModele+' (c) CEGID' ;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure EnregModele ( Mem : TMemoryStream ; Base : TLPBase ) ;
Var Combien,j,i,l,CountBase,Posit,Count : Integer ;
    Bande                               : TLPBande ;
    Ctrl                                : TControl ;
//    Mem2                                : TMemoryStream ;
Begin
Combien:=Base.ControlCount ;
Mem.Seek(0,soFromBeginning) ;
Mem.Write(ConstVModele,length(ConstVModele)) ;
posit:=mem.position ; 
Mem.Write(Combien,Sizeof(Combien)) ;
Mem.WriteComponent(Base) ;
for j:=1 to 2 do
   Begin
   CountBase:=0 ;
   For i:=0 to Base.ControlCount-1 do
       if Base.Controls[i] is TLPBande then
          Begin
          Inc(CountBase) ;
          Bande:=TLPBande(Base.Controls[i]) ;
          if j=1 then Mem.WriteComponent(Bande) else //Enregistre la Bande
             Begin //Enregistre les controls de la bande ;
             Combien:=Bande.ControlCount ;
             Posit:=Mem.Position ;
             Mem.Write(Combien,Sizeof(Combien)) ;
             Count:=0 ;
             For l:=0 to Bande.ControlCount-1 do
                 if (Bande.Controls[l] is TLPChamp) or
                    (Bande.Controls[l] is TLPImage)    then
                    Begin
                    inc(Count) ;
                    Ctrl:=TControl(Bande.Controls[l]) ;
                    Mem.WriteComponent(Ctrl) ;
                    End ;
             if Combien<>Count then
                Begin // Mise à jour du nombre de champs
                Combien:=Count ;
                Mem.Seek(Posit,soFromBeginning) ;
                Mem.Write(Combien,Sizeof(Combien)) ;
                Mem.Seek(0,soFromEnd) ;
                End ;
             End ;
          End ;
   if (J=1) and (Combien<>Countbase) then
      Begin //Mise à jour du nombre de bandes
      Combien:=CountBase ;
      Mem.Seek(Posit,soFromBeginning) ;
      Mem.Write(Combien,Sizeof(Combien)) ;
      Mem.Seek(0,soFromEnd) ;
      End ;
   End ;
CompressBinaryObjectStream(mem);
(*Mem2:=nil ;
if Mem is TBlobStream then
   Begin
   Mem2:=TMemoryStream.Create ;
   Mem2.LoadFromStream(Mem) ;
   end else
if Mem is TMemoryStream then Mem2:=TMemoryStream(Mem) ;
if assigned(mem2) then
   Begin
   CompressBinaryObjectStream(mem2); //XMG **--**
   if Mem is TBlobStream then
      Begin
      Mem.Position:=0 ;
      TStringStream(Mem).Size:=0 ; // Truncate ;
      TStringStream(Mem).CopyFrom(Mem2, 0) ;
      Mem2.Free ;
      end ;
   End ;*)
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
{$IFDEF eAGLClient}
{$ELSE}
Function EnregLP ( LPParam : TLPParam ; Base : TLPBase ) : Boolean ;
Var Q      : TQuery ;
    Mem    : TMemoryStream ;
Begin
Q:=OpenSQL('Select * from MODELES where MO_TYPE="'+LPParam.TypeEtat+'" and MO_NATURE="'+LPParam.NatEtat+'" and MO_CODE="'+LPParam.CodeEtat+'" AND MO_LANGUE="'+LPParam.Langue+'"',FALSE) ;
if Not Q.Eof then Q.Edit else Q.Insert ;
with LPParam,Q do
   Begin
   FindField('MO_TYPE').asString:=TypeEtat ;
   FindField('MO_NATURE').asString:=NatEtat ;
   FindField('MO_CODE').asString:=CodeEtat ;
   FindField('MO_LANGUE').asString:=Langue ;
   FindField('MO_LIBELLE').asString:=Libelle ;
   FindField('MO_PERSO').asString:=TrueFalseST(Personnel) ;
   FindField('MO_DEFAUT').asString:=TrueFalseST(Defaut) ;
   FindField('MO_MODELE').asString:=TrueFalseST(Modele) ;
   FindField('MO_PROTECT').asString:=TrueFalseST(Protect) ;
   FindField('MO_SUIVANT').asString:=Suivant ;
   FindField('MO_PROTECTSQL').asString:=TrueFalseST(ProtectSQL) ;
   FindField('MO_USER').asString:=User ;
   FindField('MO_DIAGTEXT').asString:=DiagText ;
   FindField('MO_FORCEDIAGTEXT').asString:=TrueFalseST(ForceDT) ;
   FindField('MO_MENU').asString:=TrueFalseST(Menu) ;
   FindField('MO_EXPORT').asString:=TrueFalseST(Exporter) ;
   End ;
Q.Post ;
Ferme(Q) ;
Q:=OpenSQL('Select * from MODEDATA where MD_CLE="'+ConstruitCle(LPParam,TRUE,FALSE)+'"',FALSE) ;
if Not Q.Eof then Q.Edit else Q.Insert ;
Q.FindField('MD_CLE').asString:=ConstruitCle(LPParam,TRUE,FALSE) ;
Mem:=TMemoryStream.Create ;
try
  Mem.seek(0,soFromBeginning) ;
  EnregModele(mem,Base) ;
  Mem.seek(0,soFromBeginning) ;
  TBlobField(Q.FindField('MD_DATA')).LoadFromStream(Mem) ;
 Finally
  Mem.Free ;
 End ;
Q.Post ;
Ferme(Q) ;

Q:=OpenSQL('Select * from MODEDATA where MD_CLE="'+ConstruitCle(LPParam,TRUE,TRUE)+'"',FALSE) ;
if Not Q.Eof then Q.Edit else Q.Insert ;
Q.FindField('MD_CLE').asString:=ConstruitCle(LPParam,TRUE,TRUE) ;
Q.FindField('MD_DATA').AsString:=LPParam.SQL ;
Q.Post ;
Ferme(Q) ;
LPParam.Nouvelle:=FALSE ;
Result:=TRUE ;
End ;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure RecupererModele( Mem : TStream ; Base : TLPBase ; pAsigneEventsChamp,pAsigneEventsBande : TProcAsigneEventsLP ; Var ChangementVersion : Boolean ) ;
Var i,j,Combien,Ps : Integer ;
    Ctrl           : TLPChamp ;
    Bande          : TLPBande ;
    Ima            : TLPImage ;
    CM             : TComponent ;
    Mem2           : TMemoryStream ;
    VImp           : Double ;
    OkVersion      : Boolean ;
Begin
ChangementVersion:=FALSE ;
Mem2:=TMemoryStream.Create ;
Mem2.CopyFrom(Mem, 0) ;
DecompressBinaryObjectStream(TMemoryStream(Mem2));
Mem.Position:=0;
Mem.CopyFrom(Mem2, 0) ;
Mem2.Free ;
Mem.Seek(0,soFromBeginning) ;
ChercheVersion(Mem,VersionModele,ConstVModele,OkVersion,Vimp) ;
ChangementVersion:=(not OkVersion) or (Vimp<>Valeur(VersionModele)) ;
Mem.Read(Combien,Sizeof(Combien)) ;
Base:=TLPBase(Mem.readComponent(Base)) ;
Base.Left:=maxintValue([8,(Base.parent.ClientWidth-Base.Width-8) div 2]) ;
Base.Top:=maxintValue([8,(Base.parent.ClientHeight-Base.Height-8) div 2]) ;
For i:=0 to Combien-1 do
  Begin // Lit les Bandes
  Bande:=TLPBande.Create(Base.Owner) ;
  Bande:=TLPBande(Mem.ReadComponent(Bande)) ;
  Bande.Name:='' ;
  Bande.Parent:=Base ;
  if assigned(pAsigneEventsBande) then pAsigneEventsBande(Bande) ;
  End ;
base.RecalculeHeight ;
for j:=0 to Base.ControlCount-1 do
  Begin
  Bande:=TLPBande(Base.Controls[j]) ;
  Mem.Read(Combien,Sizeof(Combien)) ;
  For i:=0 to Combien-1 do
     Begin //Lit les controls de la bande ;
     ps:=Mem.Position ;
     CM:=Mem.ReadComponent(nil) ;
     CM.Name:='' ; 
     if CM is TLPChamp then
        Begin
        Mem.Seek(ps,soFromBeginning) ;
        Ctrl:=TLPChamp.Create(Base.Owner) ;
        Ctrl:=TLPChamp(Mem.ReadComponent(Ctrl)) ;
        if ((Trim(Ctrl.Libelle)<>'') and ((Copy(Ctrl.Libelle,1,1)<>'#') or (Copy(Ctrl.Libelle,1,2)='##'))) or (Base.Designing) then //XMG 25/06/03
           Begin
           //Pour compatibilité
           Ctrl.Taille:=Ctrl.Taille ;
           Ctrl.Name:='' ;
           Ctrl.Parent:=Bande ;
           Ctrl.TabStop:=TRUE ;
           Ctrl.TabOrder:=Ctrl.Parent.ControlCount-1 ;
           if assigned(pAsigneEventsChamp) then pAsigneEventsChamp(Ctrl) ;
           End else Ctrl.Free ;//XMG 25/06/03
        End else
     if CM is TLPImage then
        Begin
        Mem.Seek(ps,soFromBeginning) ;
        Ima:=TLPImage.Create(Base.Owner) ;
        Ima:=TLPImage(Mem.ReadComponent(Ima)) ;
        Ima.Name:='' ;
        Ima.Parent:=Bande ;
        if assigned(pAsigneEventsChamp) then pAsigneEventsChamp(Ima) ;
        End else ;
     CM.Free ;
     End ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function ChargeModeleLP (Var LPParam : TLPParam ; Creer : Boolean ; Base : TLPBase ; pAsigneEventsChamp,pAsigneEventsBande : TProcAsigneEventsLP ; Var Err : TMC_Err ; Var ChangementVersion : Boolean ) : Boolean ;
Var Q            : TQuery ;
    Langue,Titre : String ;
    j            : TLPTypeBandes ;
    laBnd        : TLPBande ;
    TopBnd       : Integer ;
    Mem          : TStringStream  ;
Begin
Result:=FALSE ;
if Base=nil then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1002) ; Exit ; End ;
Langue:=LPParam.Langue ;
ExisteModele(LPParam.TypeEtat,LPParam.NatEtat,LPParam.CodeEtat,Langue,Not Creer) ;
Q:=OpenSQL('Select * from MODELES where MO_TYPE="'+LPParam.TypeEtat+'" and MO_NATURE="'+LPParam.NatEtat+'" and MO_CODE="'+LPParam.CodeEtat+'" AND MO_LANGUE="'+Langue+'"',TRUE) ;
if Not Q.Eof then
   Begin
   With LPParam,Q do
     Begin
     Libelle    :=FindField('MO_LIBELLE').asString ;
     Langue     :=FindField('MO_LANGUE').asString ;
     Personnel  :=TrueFalseBo(FindField('MO_PERSO').asString) ;
     Defaut     :=TrueFalseBo(FindField('MO_DEFAUT').asString) ;
     Modele     :=TrueFalseBo(FindField('MO_MODELE').asString) ;
     Protect    :=TrueFalseBo(FindField('MO_PROTECT').asString) ;
     Suivant    :=FindField('MO_SUIVANT').asString ;
     ProtectSQL :=TrueFalseBo(FindField('MO_PROTECTSQL').asString) ;
     User       :=FindField('MO_USER').asString ;
     DiagText   :=FindField('MO_DIAGTEXT').asString ;
     ForceDT    :=TrueFalseBo(FindField('MO_FORCEDIAGTEXT').asString) ;
     Menu       :=TrueFalseBo(FindField('MO_MENU').asString) ;
     Exporter   :=TrueFalseBo(FindField('MO_EXPORT').asString) ;
     Nouvelle   :=FALSE ;
     End ;
   End else
   Begin
   if not Creer then Begin Err.COde:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1003) ;  End else
      Begin
      With LPParam do
        Begin
        Libelle    :=TraduireMemoire(MC_MsgErrDefaut(1032)) ;
        Personnel  :=FALSE ;
        Defaut     :=FALSE ;
        Modele     :=FALSE ;
        Protect    :=FALSE ;
        Suivant    :='' ;
        ProtectSQL :=FALSE ;
        User       :=V_PGI.User ;
        DiagText   :='' ;
        ForceDT    :=FALSE ;
        Menu       :=FALSE ;
        Exporter   :=FALSE ;
        SQL        :='SELECT * FROM CHOIXCOD' ;
        Nouvelle   :=TRUE ;
        End ;
      End ;
   End ;
Ferme(Q) ;
if Err.Code<>0 then exit ;
Q:=OpenSQL('Select * from MODEDATA where MD_CLE="'+ConstruitCle(LPParam,FALSE,FALSE)+Langue+'"',FALSE) ;
if Not Q.Eof then
   Begin
   Mem := TStringStream.Create(Q.FindField('MD_DATA').AsString) ;
   try
     Mem.seek(0,soFromBeginning) ;
     RecupererModele(Mem,Base,pAsigneEventsChamp,pAsigneEventsBande,ChangementVersion) ;
    Finally
     Mem.Free ;
    End ;
   End else
   if not Creer then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1004) ; end else
      Begin
      Base.EnPouces:=(LPParam.TypeEtat=TypeEtatTexte) ; // //XMG 02/04/04 'T' ; //XMG 30/03/04
      Base.Rouleau:=(LPParam.TypeEtat=TypeEtatExportAscii) ; //XMG 02/04/04 'F' ; //XMG 30/03/04
      Base.Width:=800 ; //8 pouces
      Base.MargeLeft:=0 ;
      Base.MargeTop:=0 ;
      Base.MargeRight:=0 ;
      Base.MargeBottom:=0;
      Base.Units:=luChar ;
      TopBnd:=0 ;
      for j:=low(TLPTypeBandes) to High(TLPTypeBandes) do
        if not (j in [lbSubEntete,lbSubDetail,lbSubPied,lbenteterupt,lbpiedrupt]) then
          Begin
          Case j of
            lbEntete : Titre:=TraduireMemoire(MC_MsgErrDefaut(1033)) ;
            lbDetail : Titre:=TraduireMemoire(MC_MsgErrDefaut(1034)) ;
            lbPied   : Titre:=TraduireMemoire(MC_MsgErrDefaut(1035)) ;
            End ;
          LaBnd:=TLPBande.Create(Base.Owner) ;
          LaBnd.Parent:=Base ;
          LaBnd.Name:='Bande'+inttostr(ord(j)) ;
          LaBnd.height:=80 {5 Lignes} ;
          LaBnd.TypeBande:=j ;
          LaBnd.Titre:=Titre ;
          LaBnd.Top:=TopBnd ;
          TopBnd:=LaBnd.Top+LaBnd.Height ;
          if assigned(pAsigneEventsBande) then pAsigneEventsBande(LaBnd) ;
          End ;
      End ;
Ferme(q) ; //XMG 11/06/03
if Err.Code<>0 then exit ;
Q:=OpenSQL('Select * from MODEDATA where MD_CLE="'+ConstruitCle(LPParam,FALSE,TRUE)+Langue+'"',TRUE) ;
if Not Q.Eof then LPParam.SQL:=Q.FindField('MD_DATA').AsString
   else if Not Creer then Begin Err.Code:=-1 ; Err.Libelle:=MC_MsgErrDefaut(1005) ; end ;
Ferme(Q) ;
Result:=(Err.Code=0) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
Const VersionExpt  = '2.00' ; //1.00 original ;; 2.00 modèles compressées
      ConstVersion = 'Modele impresion texte Version '+VersionExpt+' (c) CEGID' ;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function LP_ImporteModeleInterne ( NomFichier : String ; Var LPParam : TLPParam ; Base : TLPBase ; pAsigneEventsChamp,pAsigneEventsBande : TProcAsigneEventsLP ; Var Err : TMC_Err ; Var ChangementVersion : Boolean ) : Boolean ;
Var Imp                : TMemoryStream ;
    Modele             : TStringStream ;
    St,TypEtat,NatEtat : String ;
    VImp               : Double ;
    OkVersion          : Boolean ;
Begin
Result:=FALSE ;
err.code:=0 ; err.Libelle:='' ;
TypEtat:='' ;
NatEtat:='' ;
if trim(LPParam.TypeEtat)<>'' then TypEtat:=LPParam.TypeEtat ;
if trim(LPParam.NatEtat)<>'' then NatEtat:=LPParam.NatEtat ;
Imp:=TMemoryStream.Create ;
Modele:=TStringStream.Create('') ;
try
  Imp.LoadFromFile(Nomfichier) ;
  Imp.Seek(0,soFromBeginning) ;
  ChercheVersion(imp,VersionExpt,ConstVersion,OkVersion,Vimp) ;
  if (OkVersion) then
     Begin
     if VImp<=Valeur(VersionExpt) then
        Begin
        if VImp=2.00 then LpParam:=TLPParam(Imp.ReadComponent(LPParam)) ;
        Modele.CopyFrom(Imp,Imp.Size-Imp.Position) ;
        RecupererModele(Modele,Base,pAsigneEventsChamp,pAsigneEventsBande,ChangementVersion) ;
        if VImp=1.00 then LpParam:=TLPParam(Modele.ReadComponent(LPParam)) ;
        if trim(TypEtat)<>'' then LPParam.TypeEtat:=TypEtat ;
        if trim(NatEtat)<>'' then LPParam.NatEtat :=NatEtat ;
        LPParam.Nouvelle:=TRUE ;
        Result:=TRUE ;
        End else St:='V.Imp.-'+floattostr(VImp)+' / V.Base.-'+VersionExpt ;
     End ;
  if (Not OkVersion) or (VImp>Valeur(VersionExpt)) then
     Begin
     Err.Code:=-1 ;
     Err.Libelle:=MC_MsgErrDefaut(1006)+' '+St ;
     End ;
 except
  Err.Code:=-1 ; Err.Libelle:=Format(MC_MsgErrDefaut(1007),[ExtractFileName(NomFichier)]) ;
  if V_PGI.SAV then raise ;
 End ;
Imp.Free ;
Modele.Free ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
{$IFDEF eAGLClient}       //YCP 19/05/03
{$ELSE}
Function LP_ImporteModele ( NomFichier : String ; Var Err : TMC_Err ; Var ChangementVersion : Boolean ) : Boolean ;
Var FForm   : TForm ;
    LPParam : TLPParam ;
    Base    : TLPBase ;
Begin
err.code:=0 ; err.Libelle:='' ; 
FForm:=TForm.Create(Application) ;
try
  LPParam:=TLPParam.create(FForm) ;
  Base:=TLPBase.Create(FForm) ;
  Base.Parent:=FForm ;
  if LP_importeModeleInterne(NomFichier,LPParam,Base,nil,nil,Err,ChangementVersion) then
     Begin
     LPParam.User       :=V_PGI.User ;
     _BeginTrans ;
     try
        if EnregLP(LPParam,Base) then _CommitTrans else
           Begin
           Err.Code:=1 ;
           abort ; //Force le rollback ;
           End ;
       except
        _RollBack ;
       End ;
     End ;
 Finally
  FForm.Free ;
 End ;
if Err.Code>0 then Err.Libelle:=MC_MsgErrDefaut(Err.code) ;
Result:=(Err.Code=0) ;
End ;
{$ENDIF eAGLClient}

//////////////////////////////////////////////////////////////////////////////////
Function LP_ExporteModeleInterne ( NomFichier : String ; LPParam : TLPParam ; Base : TLPBase ; Var Err : TMC_Err ) : Boolean ;
Var Exp    : TMemoryStream ;
    Modele : TMemoryStream ;
Begin
Result:=FALSE ;
err.code:=0 ; err.Libelle:='' ;
Exp:=TMemoryStream.Create ;
Modele:=TMemoryStream.Create ;
try
  Exp.Seek(0,soFromBeginning) ;
  Exp.Write(ConstVersion,length(ConstVersion)) ;
  Exp.WriteComponent(LPParam) ;
  EnregModele(Modele,Base) ;
  Exp.CopyFrom(Modele,0) ;
  Exp.SaveToFile(Nomfichier) ;
  Result:=TRUE ;
 except
  Err.Code:=-1 ; Err.Libelle:=Format(MC_MsgErrDefaut(1008),[ExtractFileName(NomFichier)]) ;
 End ;
Exp.Free ;
Modele.Free ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function LP_ExporteModele ( NomFichier : String ; TpEtat,NatureEtat,Code,Lng : String ; Var Err : TMC_Err ) : Boolean ;
Var FForm   : TForm ;
    LPParam : TLPParam ;
    Base    : TLPBase ;
    ChangementVersion : Boolean ;
Begin
err.code:=0 ; err.Libelle:='' ;
FForm:=TForm.Create(Application) ;
try
  LPParam:=TLPParam.create(FForm) ;
  Base:=TLPBase.Create(FForm) ;
  Base.Parent:=FForm ;
  With LPParam do
    Begin
    TypeEtat:=TpEtat ;
    NatEtat:=NatureEtat ;
    CodeEtat:=Code ;
    Langue:=Lng ;
    End ;
  If ChargeModeleLP(LPParam,FALSE,Base,nil,nil,Err,ChangementVersion) then LP_ExporteModeleInterne(Nomfichier,LPParam,base,Err)
 Finally
  FForm.Free ;
 End ;
if Err.Code>0 then Err.Libelle:=MC_MsgErrDefaut(Err.code) ;  
Result:=(Err.COde=0) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Constructor TLPPrinter.Create ( AOwner : TComponent ; AImprimante : String ) ;
Begin
Inherited Create(AOwner) ;
FPrinterInterne:=nil ;
if AIMprimante='EXP' then FPrinterInterne:=TLP_EXPORTERFICHIER.Create(AOwner,AImprimante) else //XMG 30/03/04
if AIMprimante='001' then FPrinterInterne:=TLP_EPSONTM210.Create(AOwner,AImprimante) else
if AIMprimante='002' then FPrinterInterne:=TLP_EPSONTM675.Create(AOwner,AImprimante) else
if AIMprimante='003' then FPrinterInterne:=TLP_EPSONTM950.Create(AOwner,AImprimante) else
if AIMprimante='004' then FPrinterInterne:=TLP_EPSONTM5000.Create(AOwner,AImprimante) else
if AIMprimante='005' then FPrinterInterne:=TLP_EPSONTM88II.Create(AOwner,AImprimante) else
if AIMprimante='006' then FPrinterInterne:=TLP_EPSONTM300.Create(AOwner,AImprimante) else
if AIMprimante='007' then FPrinterInterne:=TLP_EPSONTM88II.Create(AOwner,AImprimante) else
if AIMprimante='008' then FPrinterInterne:=TLP_EPSONTM6000.Create(AOwner,AImprimante) else ;
if AIMprimante='009' then FPrinterInterne:=TLP_STARTSP700.Create(AOwner,AImprimante) else ;

if not assigned(FPrinterInterne) then Raise EComponentError.CreateFmt(MC_MsgErrDefaut(1061),[AImprimante]) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
destructor TLPPrinter.Destroy ;
Begin
if assigned(FPrinterInterne) then FPrinterInterne.Free ;
Inherited destroy ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetLastError : TImprimErreurs ;
Begin
Result:=ieUnknown ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.Lasterror ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetLargeur : Integer ;
Begin
Result:=-1 ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.Largeur ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.SetLargeur ( Value : Integer ) ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.Largeur:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetPort : TComPortNumber ;
Begin
result:=pnNone ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.Port ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.SetPort (AValue : TComPortNumber ) ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.Port:=AValue ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.ChargePortEtParams(Aport,AParams : String ; Var err : TMC_Err ) : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.ChargePortEtParams(APort,Aparams,Err) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.isConnected : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.IsConnected ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetSqce (Indice : TLP_SqcesPossibles ; Activer : Boolean ) : String ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.Sqce[Indice,Activer] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetInitialise : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.Initialise ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.SetInitialise   ( Value : Boolean ) ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.Initialise:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetSqceInUse (Indice : TLP_SqcesPossibles) : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.SqceInUse[Indice] ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.SetSqceInUse (Indice : TLP_SqcesPossibles ; Value : Boolean ) ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.SqceInUse[Indice]:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.SetInTestMateriel ( Value : Boolean ) ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.InTestMateriel:=Value ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetInTestMateriel : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.InTestMateriel ;
End ;
///////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.InitLastligne ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.LastLigne:=-1 ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  TLPPrinter.GetLastLigne : Integer  ;
Begin
REsult:=-1 ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.LastLigne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.SetLastLigne ( Value : Integer ) ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.LastLigne:=Value ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.OuvrirImprimante ( Var Err : TMC_Err ) : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.OuvrirImprimante(Err) ;
if (not result) and (Err.Code=0) and (Lasterror<>ieError) then Begin Err.Code:=-1 ; Err.Libelle:=Format(MC_MsgErrDefaut(1024),[FPrinterInterne.LastErrorW,SysErrorMessage(FPrinterInterne.LastErrorW)]) ; End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.FermeImprimante( Var Err : TMC_Err ) : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.FermeImprimante(Err) ;
if (not result) and (Err.Code=0) then Begin Err.Code:=-1 ; Err.Libelle:=Format(MC_MsgErrDefaut(1024),[FPrinterInterne.LastErrorW,SysErrorMessage(FPrinterInterne.LastErrorW)]) ; End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.InitImprimante ( Var Err : TMC_Err ) : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.InitImprimante(Err) ;
if (not result) and (Err.Code=0) then Begin Err.Code:=-1 ; Err.Libelle:=Format(MC_MsgErrDefaut(1024),[FPrinterInterne.LastErrorW,SysErrorMessage(FPrinterInterne.LastErrorW)]) ; End ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.InitNewLigne ;
Begin
if Not Assigned(FPrinterInterne) then exit ;
FPrinterInterne.InitNewLigne ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.SautPage : Boolean ;
Begin
Result:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.SautPage ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.Demarre ;
BEgin
If Assigned(FPrinterInterne) then FPrinterInterne.Demarre ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.DoTranslate ( St : String ) : String ; //XMG 30/03/04
Begin
Result:='' ;
if Not Assigned(FPrinterInterne) then exit ;
REsult:=FPrinterInterne.FaireTranslate(St) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TLPPrinter.FormatPolice ( zn : TLP_ZoneImp) ; //Seq : TLP_SetSqcesPossibles ; Colonne, Longeur : Integer ; Extra : String ) ;
Begin
if Assigned(FPrinterInterne) then FPrinterInterne.FormatPolice(Zn) ; //Seq,Colonne,Longeur,Extra)
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TLPPrinter.EcrireLigne(St : String ; Ligne : Integer ) : Boolean ;
Begin
REsult:=FALSE ;
if Not Assigned(FPrinterInterne) then exit ;
Result:=FPrinterInterne.EcrireLigne(St,Ligne) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function ExisteModele(TypeEtat,NatEtat,ModeleEtat : String ; Var Langue :String ; OkLangueDef : Boolean  ) : Boolean ;
//on a repris la truc de EAGL, pour à ne pas trop compliquer.... //XMG
Var i : Integer ;
    LangOrg : String ;
Begin
Result:=FALSE ; i:=0 ;
LangOrg:=Langue ;
while (((i<3) and (OkLangueDef)) or (i<2)) and (Not Result) do
  Begin
  inc(i) ;
  if (trim(Langue)='') or ((i>1) and (OkLangueDef)) then
     Begin
     if i<3 then begin Langue:=V_PGI.LanguePrinc ; Inc(i,ord(i=1)) ; end
        else Langue:='FRA'
     End ;
  if trim(Langue)<>'' then
     Begin
     Result:=ExisteSQL('Select (1) from MODELES where MO_TYPE="'+TypeEtat+'" and MO_NATURE="'+NatEtat+
                                               '" and MO_CODE="'+ModeleEtat+'" AND MO_LANGUE="'+Langue+'"') ;
     End ;
  End ;
if not Result then Langue:=LangOrg ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
end.

