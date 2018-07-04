{***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 20/02/2007
Modifié le ... :   /  /
Mots clefs ... :
*****************************************************************}
unit FPRAPPROCRE_TOF;

interface

uses
   SysUtils
   , Classes
   , Graphics
   , Controls
   , Forms
   , Dialogs
   , Grids
   , Hctrls
   , HEnt1
   , windows
   , HPanel
   , Uiutil
   , HSysMenu
   , HTB97
   , HStatus
   , UTob
   , ed_tools
   , HMsgBox
   , Mask
   , StdCtrls
   , ExtCtrls
   , Math  // RoundTo
   , uTof,
{$IFDEF EAGLCLIENT}
    MaineAGL,
{$ELSE}
    db,
    fe_main,
{$ENDIF}
    Saisie,
    SaisBor,
    ComCtrls,
    Menus,
    HDB,
    HQry,
    Vierge,
    UObjFiltres,
    Ent1,
    SaisUtil,
    UtilSais,
    DelVisuE,
    AGLInit,
    UtilSoc,
    paramsoc,
   {$IFDEF eAGLCLient}
    UtileAGL,
    uObjEtats            // TObjEtats.GenereEtatGrille
   {$ELSE}
    PrintDBG,
   {$IFNDEF DBXPRESS}
    dbtables
   {$ELSE}
    uDbxDataSet
   {$ENDIF}
   {$ENDIF eAGLCLient}
   ;


const COL_COMPTE=0;
      COL_LIBELLE=1;
      COL_SOLDECRE=2;
      COL_SOLDECOMPTA=3;
      COL_ECART=4;
      COL_TYPECOMPTE=5;

type
  TParamEcr = class
    Journal : string;
    Libelle : string;
    Date : TDateTime;
    DateCalcul : TDateTime;
  end;

  TLigneEmprunt = class
    CodeEmprunt : String;
    Compte : string;
    TypeCompte : String;
    Libelle : string;
    SoldeCRE : double;
    DebitECR : double;
    CreditECR : double;
    Date : TDateTime;
    Sens : String ;
  end;

  TOF_FPRAPPROCRE = class(TOF)
    PBouton: TToolWindow97;
    Dock: TDock97;
    HMTrad: THSystemMenu;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BZoomDetail: TToolbarButton97;
    PDATERAPPRO: THPanel;
    HLabel1: THLabel;
    DATERAPPROCHEMENT: THCritMaskEdit;
    BCALCUL: TToolbarButton97;
    Type_ecr: TLabel;
    FEcran   : TFVierge ;
    DateRappro : THEdit ;
    CECRN : TCheckBox ;
    CECRS : TCheckBox ;
    CECRU : TCheckBox ;
    procedure OnArgument(S: string); override;
    procedure OnZoomEcritureClick(Sender: TOBject);
    procedure OnZoomEmpruntClick(Sender: TOBject);
    procedure OnClose; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnDisplay; override;
    procedure OnCancel; override;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure FListeDblClick(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure BCalculClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure OnKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées}
 
    FListe : THGrid ;

    procedure VideListe(L : TList);
    procedure CalculCREECR (ListeE : TList) ;
    procedure RecupMtEcri (ListeE : TList) ;
    procedure AfficheListeEcriture (ListeE : TList) ;
    procedure MajListe(fListeE: TList;Compte,Libelle: string; CRESolde,ECRDebit,ECRCredit: double;Sens: string) ;
    procedure CreaListeEmprunt ( var L : TList;Code : string);
    procedure Rech_MtPrecis(Code, TypeGene : String ; date_fin : TDateTime ; var Mt_solde, Mt_inter, Mt_assur : Double);
    procedure MajLigneCRE (ChpCompte, Compte, Emprunt, TypeGene : String ; date_fin : TDateTime ; Montant : Double ; var L:TList);
    procedure MajLigneRegul (ChpCompte, TypeGene, Compte, Emprunt : String ; var L:TList);

  public
    { Déclarations publiques}
  end;


procedure EtatRapproCRECompta;

implementation

uses
  ZEcriMvt_TOF
  , CalcOLE
//  , AMLISTE_TOF
  , UTofConsEcr,
  TofMulEmprunt;

procedure EtatRapproCRECompta;
begin
 AGLLanceFiche('FP','FPRAPPROCRE','','',';;');
end;

function CompareCompte (Item1,Item2:Pointer) : integer;
var Ecr1,Ecr2 : TLigneEmprunt;
begin
  Ecr1 := Item1;Ecr2 := Item2;
  if Ecr1.Compte > Ecr2.Compte then Result := 1
  else if Ecr1.Compte < Ecr2.Compte then Result := -1
  else Result := 0;
end;


procedure TOF_FPRAPPROCRE.OnArgument(S: string);
begin
inherited ;

  // Récup interface
  FEcran := TFVierge(Ecran) ;

  FListe  := THGrid( GetControl('FListe',True) ) ;

  DateRappro := THEdit(GetControl('DATERAPPRO')) ;
  DateRappro.Text := DateToStr(VH^.Encours.Fin) ;

  TToolBarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;

  TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCalculClick;
  TToolBarButton97(GetControl('BVALIDER')).OnClick := BCalculClick;
  FListe.OnDblClick := FListeDblClick;

  CECRN := TCheckBox(GetControl('CECRN')) ;
  CECRS := TCheckBox(GetControl('CECRS')) ;
  CECRU := TCheckBox(GetControl('CECRU')) ;

  Ecran.OnKeyDown := FormKeyDown;

  TToolBarButton97(GetControl('BZOOM')).OnClick := OnZoomEcritureClick;
  TToolBarButton97(GetControl('BZOOMEMPRUNT')).OnClick := OnZoomEmpruntClick;

  {Format cellules montants}
  fListe.ColFormats[COL_SOLDECRE]:='#,##0.00' ;
  fListe.ColAligns[COL_SOLDECRE]:=taRightJustify ;
  fListe.ColFormats[COL_SOLDECOMPTA]:='#,##0.00' ;
  fListe.ColAligns[COL_SOLDECOMPTA]:=taRightJustify ;
  fListe.ColFormats[COL_ECART]:='#,##0.00' ;
  fListe.ColAligns[COL_ECART]:=taRightJustify ;

end;

procedure TOF_FPRAPPROCRE.OnZoomEcritureClick(Sender: TOBject);
begin
  FListeDblClick(nil);
end;

procedure TOF_FPRAPPROCRE.OnZoomEmpruntClick(Sender: TOBject);
var
  Compte,TypeCompte: string;
begin
  Compte := FListe.Cells[COL_COMPTE,Fliste.Row];
  if Compte = '' then exit;
  {On va se placer sur le Mul Emprunt en renseignant le champ du code de l'emprunt.
  Il se trouve que le compte peut aussi être l'un des 4 autres comptes (charges d'avance,...)
  Si le compte correspond à un seul type de compte <> numero d'emprunt, on se positionne
  sur celui-là (un même compte pouvant servir à 2 choses différentes selon l'emprunt)}
  TypeCompte := FListe.Cells[COL_TYPECOMPTE,Fliste.Row];
  CPLanceFiche_Emprunt(Compte+';'+TypeCompte);
end;

procedure TOF_FPRAPPROCRE.OnClose ;
begin
  inherited ;
end;

procedure TOF_FPRAPPROCRE.CalculCREECR(ListeE : TList) ;
begin

  SourisSablier;
  InitMove (1,'');
  CreaListeEmprunt(ListeE,'');
  RecupMtEcri (ListeE) ;
  FiniMove;

end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Recherche des montants de régularisation
Mots clefs ... :
*****************************************************************}
procedure TOF_FPRAPPROCRE.Rech_MtPrecis(Code, TypeGene : String ; date_fin : TDateTime ; var Mt_solde, Mt_inter, Mt_assur : Double);
var
   lStSQL : String ;
   lTbEcheances, lTbDetail : Tob ;
   lEche : Integer ;
   Date_eche, Date_decal : TDateTime ;
   Periode, Quota : Double ;
   QP : TQuery ;
begin

    // Extraction des données de l'emprunt
    lStSQL := 'SELECT ECH_DATE FROM FEMPRUNT, FECHEANCE';
    lStSQL := lStSQL + ' WHERE ECH_CODEEMPRUNT = EMP_CODEEMPRUNT';
    lStSQL := lStSQL + ' AND EMP_CODEEMPRUNT ="'+Code+ '"';
    lStSQL := lStSQL + ' AND EMP_DATEDEBUT <="'+USDateTime(Date_fin)+ '"';
    lStSQL := lStSQL + ' ORDER BY ECH_CODEEMPRUNT, ECH_DATE';

    lTbEcheances := Tob.Create('',Nil,-1);

    lTbEcheances.LoadDetailFromSQL(lStSQL);

    Date_Decal:=VH^.Encours.Deb;
    Date_Eche:= VH^.Encours.Deb;

    Mt_solde:=0;
    Mt_inter:=0;
    Mt_assur:=0;
    Quota:=1;
    
    // Boucle sur les échéances
    For lEche:= 0 to lTbEcheances.Detail.Count- 1 do
    begin

       lTbDetail := lTbEcheances.Detail[lEche];
       If TypeGene<>'R' then
       begin
         If lTbDetail.GetDateTime('ECH_DATE')<=Date_fin then
            date_eche:= lTbDetail.GetDateTime('ECH_DATE');
         If lTbDetail.GetDateTime('ECH_DATE')> Date_fin then
         begin
            date_decal:= lTbDetail.GetDateTime('ECH_DATE');
            break;
         end;
       end
       else
       begin
         If lTbDetail.GetDateTime('ECH_DATE')< Date_fin then
            date_decal:= lTbDetail.GetDateTime('ECH_DATE');
         If lTbDetail.GetDateTime('ECH_DATE')>=Date_fin then
         begin
            date_eche:= lTbDetail.GetDateTime('ECH_DATE');
            break;
         end;
       end;

    end;

    lTbEcheances.Free;

    If date_eche<>date_decal then //25.06.07  Sinon pas d'échéances
    begin
      Periode:=1;
      If TypeGene='A' then {Avance}
      begin // On enlève la partie qu'on a déjà payée, de la date choisie jusqu'à la prochaine échéance
        Periode:=Date_decal-Date_eche;
        Quota:=Date_decal-Date_fin ; {30.05.2008 erreur de un jour}
      end
      else     {Retard}
      If TypeGene='R' then
      begin // On enlève la partie qu'on n'a pas encore payée, de la précédente échéance à la date choisie
        Periode:=Date_eche-Date_decal;
        Quota:=Date_fin-Date_decal;
      end;

      {Calcul Montant intérêts & assurance}
      QP := OpenSQL ('SELECT ECH_SOLDE, ECH_INTERET, ECH_ASSURANCE'
                    +' FROM FECHEANCE WHERE ECH_DATE="'+UsDateTime(Date_eche)
                    +'" AND ECH_CODEEMPRUNT ="'+Code
                    +'"', TRUE);
      if not QP.EOF then
      begin
        Mt_solde := QP.FindField('ECH_SOLDE').AsFloat ;
        Mt_inter := QP.FindField('ECH_INTERET').AsFloat ;
        Mt_assur := QP.FindField('ECH_ASSURANCE').AsFloat ;
      end ;
      Ferme (QP);

      {Calcul du prorata}
      Mt_inter:=Mt_inter*Quota/Periode ;
      Mt_assur:=Mt_assur*Quota/Periode ;
    end;

end;

procedure TOF_FPRAPPROCRE.CreaListeEmprunt ( var L : TList;Code : string);
var
QEmp, Q : TQuery;
sSelect , sWhere, sOrder, emprunt, typegene : String;
DateDeb, DateRap : TDateTime ;
lInJour, lInMois, lInAn : Word ;
begin

   QEmp := OpenSQL ('SELECT EMP_CODEEMPRUNT, EMP_NUMCOMPTE, EMP_CPTFRAISFINA, EMP_DATECONTRAT, EMP_DATEDEBUT, '
      +'EMP_CPTASSURANCE, EMP_CPTCHARGAVAN, EMP_CPTINTERCOUR FROM FEMPRUNT WHERE EMP_STATUT<>4 ORDER BY EMP_CODEEMPRUNT', TRUE);  //WHERE EMP_CODEEMPRUNT = "3"

   while not QEmp.Eof do
   begin

      DateRap:=StrToDate(DateRappro.Text);
      DecodeDate(DateRap,lInAn,lInMois,lInJour);
      DateDeb:=VH^.Encours.Deb; {FQ20853  27.06.07  YMO}

      Emprunt:=QEmp.FindField('EMP_CODEEMPRUNT').AsString;

      sSelect :='SELECT SUM(ECH_SOLDE) SOLDE, '
      +'SUM(ECH_INTERET) INTERET, SUM(ECH_ASSURANCE) ASSURANCE '
      +'FROM FECHEANCE ' ;

      sWhere:=' WHERE ECH_CODEEMPRUNT="'+Emprunt+'" '
        + ' AND ECH_DATE<="'+USDateTime(DateRap)+'"'
        + ' AND ECH_DATE>="'+USDateTime(DateDeb)+'"';

      Q := OpenSQL (sSelect+sWhere+sOrder, TRUE);

      try
        {FQ18684 YMO 25/04/07 Détermination du type de régularisation}
        {FQ22434 YMO 18.02.08 Mauvais format pour comparatif de dates}
        If QEmp.FindField('EMP_DATECONTRAT').AsDateTime < QEmp.FindField('EMP_DATEDEBUT').AsDateTime then
          TypeGene:='R'  {IC}
        else
          TypeGene:='A'; {CCA}

        {Alimentation du cumul pour les Intérêts et l'Assurance}
        MajLigneCRE ('EMP_CPTFRAISFINA',QEmp.FindField('EMP_CPTFRAISFINA').AsString, Emprunt, TypeGene, DateRap, Q.FindField('INTERET').AsFloat, L);
        MajLigneCRE ('EMP_CPTASSURANCE',QEmp.FindField('EMP_CPTASSURANCE').AsString, Emprunt, TypeGene, DateRap, Q.FindField('ASSURANCE').AsFloat, L);

        {Solde issu de la fiche d'emprunt à la date rappro}
        MajLigneRegul ('EMP_NUMCOMPTE','S',QEmp.FindField('EMP_NUMCOMPTE').AsString, Emprunt, L);

        {Calcul des régularisations par rapport à la date de référence}


        If TypeGene='R' then {IC}
          MajLigneRegul ('EMP_CPTINTERCOUR', TypeGene, QEmp.FindField('EMP_CPTINTERCOUR').AsString, Emprunt, L)
        else                 {CCA}  {FQ21663  14.12/2007  YMO Prise en comptes des charges constatées d'avance}
          MajLigneRegul ('EMP_CPTCHARGAVAN', TypeGene, QEmp.FindField('EMP_CPTCHARGAVAN').AsString, Emprunt, L);



      finally
        Ferme (Q);
        FiniMove;
      end ;

    QEmp.Next;
   end;

   Ferme (QEmp);

end;

function RechercheCompte(L : TList; Compte : String) : TLigneEmprunt;
var
  MyRecord : TLigneEmprunt;
  i : integer;
begin
  Result := nil;
  for i := 0 to L.Count-1 do
  begin
    MyRecord := L.Items[i];
    if MyRecord.Compte=Compte then
    begin
      Result := MyRecord;
      break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Mise à jour par compte du total du montant de régularisation
Suite ........ : TypeGene = (A)vance ou (R)etard (intérêts courus)
Mots clefs ... :
*****************************************************************}
procedure TOF_FPRAPPROCRE.MajLigneRegul (ChpCompte, TypeGene, Compte, Emprunt : String ; var L:TList);
var
  MyRecord : TLigneEmprunt;
  Montant : Double ;
  Mt_solde, Mt_inter, Mt_assur : Double ;
begin

//  MyRecord := nil;

  if (Compte='') Or (not Presence('GENERAUX', 'G_GENERAL', Compte)) then
      exit ;

  Rech_MtPrecis(Emprunt, TypeGene, StrToDate(DateRappro.Text), Mt_solde, Mt_inter, Mt_assur);
  If TypeGene ='S' then
    Montant:=Mt_solde
  else
    Montant:=Mt_inter+Mt_assur;

  MyRecord:=RechercheCompte(L, Compte) ;

  if (MyRecord = nil) then
  begin
    MyRecord := TLigneEmprunt.Create;
    MyRecord.Compte:=Compte ;
    MyRecord.SoldeCRE:=0 ;

    L.Add(MyRecord);
  end;
  {FQ20003 YMO 24/04/07}
  If (MyRecord.TypeCompte= '') Or (ChpCompte= 'EMP_NUMCOMPTE') then
      MyRecord.TypeCompte:=ChpCompte;

  MyRecord.SoldeCRE  := MyRecord.SoldeCRE+ RoundTo(Montant, -2);

end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /
Description .. : Mise à jour par compte du total du montant de l'emprunt
Mots clefs ... :
*****************************************************************}
procedure TOF_FPRAPPROCRE.MajLigneCRE (ChpCompte, Compte, Emprunt, TypeGene : String ; date_fin : TDateTime ; Montant : Double ; var L:TList);
var
  MyRecord : TLigneEmprunt;
  Mt, Mt_N1, Mt_N : double;
  lInJour, lInMois, lInAn : Word;
  Date_Deb : TDateTime;
begin

  //MyRecord := nil;

  if (Compte='') Or (not Presence('GENERAUX', 'G_GENERAL', Compte)) then
     exit ;

  DecodeDate(Date_fin,lInAn,lInMois,lInJour);

  {FQ20853  27.06.07  YMO}
  Date_Deb:=VH^.Encours.Deb - 1;

  If ChpCompte='EMP_CPTFRAISFINA' then  {Interets}
  begin
    Rech_MtPrecis(Emprunt, TypeGene, Date_Deb, Mt, Mt_N1, Mt);
    Rech_MtPrecis(Emprunt, TypeGene, Date_Fin, Mt, Mt_N, Mt);
  end
  else
  begin                                 {Assurance}
    Rech_MtPrecis(Emprunt, TypeGene, Date_Deb, Mt, Mt, Mt_N1);
    Rech_MtPrecis(Emprunt, TypeGene, Date_Fin, Mt, Mt, Mt_N);
  end;

  {FQ20853  27.06.07  YMO Charges réelles}
  If TypeGene='A' then
    Montant:=Montant+Mt_N1-Mt_N
  else                            {Mt+CCAN1-CCAN-ICN1+ICN}
    Montant:=Montant-Mt_N1+Mt_N;

  MyRecord:=RechercheCompte(L, Compte) ;

  if (MyRecord = nil) then
  begin
    MyRecord := TLigneEmprunt.Create;
    MyRecord.Compte:=Compte ;
    MyRecord.SoldeCRE:=0 ;

    L.Add(MyRecord);
  end;

  {FQ20003 YMO 24/04/07}
  If (MyRecord.TypeCompte= '') Or (ChpCompte= 'EMP_NUMCOMPTE') then
      MyRecord.TypeCompte:=ChpCompte;

  MyRecord.SoldeCRE  :=  MyRecord.SoldeCRE+ RoundTo(Montant, -2);

end;


////////////////////////////////////////////////////////////////////////////////
procedure TOF_FPRAPPROCRE.VideListe(L : TList);
var i:integer;
    MyRecord:TLigneEmprunt;
begin
  if L=Nil then Exit ; if L.Count<=0 then Exit ;
  for i:=0 to (L.Count - 1) do
  begin
    MyRecord:=L.Items[i];
    MyRecord.Free ;
  end;
  L.Clear ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /    
Description .. : Recupération des montants des ecritures
Mots clefs ... : 
*****************************************************************}
procedure TOF_FPRAPPROCRE.RecupMtEcri (ListeE : TList) ;
var
  MyRecord : TLigneEmprunt;
  Q : TQuery;
  LaTob,TGene: Tob ;
  i: integer ;
  StChamps,wSens: string ;
  MtDebECR,MtCreECR: double ;
  TCumul : TabloExt;
  DateRap, DateDeb : TDateTime;
  Typeecr, Lib, Exo : string;
begin
  LaTob:=Tob.Create('GENERAUX',nil,-1) ;
  DateRap := StrToDate(DATERAPPRO.Text);
  exo:=VH^.EnCours.Code;
  DateDeb:=VH^.EnCours.Deb;
  StChamps:='G_GENERAL, G_LIBELLE, G_SENS ' ;
  try

    if CECRN.Checked then typeEcr:=typeEcr+'N';
    if CECRS.Checked then typeEcr:=typeEcr+'S';
    if CECRU.Checked then typeEcr:=typeEcr+'U';

    for i:=0 to ListeE.Count-1 do
    begin
      MyRecord:=ListeE.Items[i] ;
      if MyRecord.Compte='' then continue ;

      Q:=OpenSQL ('SELECT '+StChamps+' FROM GENERAUX WHERE G_GENERAL="'+MyRecord.Compte+'"', TRUE);

      LaTob.LoadDetailDB('GENERAUX','','',Q,false,true) ;
      ferme(Q) ;

      If LaTob=Nil then continue;

      TGene:=LaTob.Detail[0] ;
      wSens :=TGene.GetString('G_SENS') ;
      Lib:=TGene.GetString('G_LIBELLE') ;

      GetCumul('GEN',MyRecord.Compte,'',typeecr,'','',Exo,DateDeb,DateRap,False,False,nil,TCumul,False);

      MtCreECR := TCumul[2];
      MtDebECR := TCumul[1];

      MajListe(ListeE,MyRecord.Compte,TGene.GetString('G_LIBELLE'),0,MtDebECR,MtCreECR,wSens) ;
    end;

  finally
    LaTob.Free ;
  end ;
end;

procedure TOF_FPRAPPROCRE.MajListe(fListeE: TList;Compte,Libelle: string; CRESolde,ECRDebit,ECRCredit: double;Sens: string) ;
var i: integer ; bTrouve: boolean ; MyRecord: TLigneEmprunt ;
begin
  bTrouve := false; i:=0 ;
  while (i<=fListeE.Count-1) and not bTrouve do
  begin
    bTrouve:=(TLigneEmprunt(fListeE.Items[i]).Compte=Compte) ;
    if not bTrouve then inc(i);
  end;
  if bTrouve then
  begin
    MyRecord:=TLigneEmprunt(fListeE.Items[i]) ;
    MyRecord.DebitECR :=MyRecord.DebitECR+ECRDebit;
    MyRecord.CreditECR :=MyRecord.CreditECR+ECRCredit;
    MyRecord.SoldeCRE :=MyRecord.SoldeCRE+CRESolde;
    if MyRecord.Libelle='' then MyRecord.Libelle:=Libelle ;
    if Sens<>''           then MyRecord.Sens:=Sens ;
  end
  else
  begin
      MyRecord        :=TLigneEmprunt.Create ;
      MyRecord.Compte :=Compte ;
      MyRecord.Libelle:=Libelle ;
      MyRecord.DebitECR :=ECRDebit ;
      MyRecord.CreditECR :=ECRCredit ;
      MyRecord.SoldeCRE :=CRESolde ;
      if Sens<>'' then MyRecord.Sens:=Sens ;
      fListeE.Add(MyRecord);
  end;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :   /  /    
Description .. : Affichage des montants dans la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_FPRAPPROCRE.AfficheListeEcriture (ListeE : TList) ;
var
  i, j : integer;
  MyRecord : TLigneEmprunt;
  SoldeCRE, SoldeCpta, SoldeDiff : double;
begin
  j:=0;
  ListeE.Sort(CompareCompte);

  FListe.RowCount := ListeE.Count+1;
  for i:=0 to ListeE.Count-1 do
  begin
    MyRecord:=ListeE.Items[i];

    SoldeCpta:=Abs(MyRecord.CreditECR-MyRecord.DebitECR) ; {25.06.07  FQ20837 YMO pas de signe sur le solde compta}
    SoldeCRE :=MyRecord.SoldeCRE ;
    SoldeDiff:=SoldeCpta-SoldeCRE ;

    inc(j);

    fListe.Cells[COL_COMPTE,j] :=MyRecord.Compte;
    fListe.Cells[COL_TYPECOMPTE,j] :=MyRecord.TypeCompte;
    fListe.Cells[COL_LIBELLE,j]:=MyRecord.Libelle;

    fListe.Cells[COL_SOLDECRE,j]  :=StrfMontant(SoldeCRE,13,V_PGI.OkDecV,'',True);
    fListe.Cells[COL_SOLDECOMPTA,j]:=StrfMontant(SoldeCpta,13,V_PGI.OkDecV,'',True);
    fListe.Cells[COL_ECART,j]      :=StrfMontant(SoldeDiff,13,V_PGI.OkDecV,'',True);
    
  end;
end;


procedure TOF_FPRAPPROCRE.FListeDblClick(Sender: TObject);
var Compte: string;
begin
  Compte := FListe.Cells[COL_COMPTE,Fliste.Row];
  if Compte = '' then exit;

  OperationsSurComptes(Compte,DateToStr(VH^.Encours.Fin), '', '',false) ;

end;



procedure TOF_FPRAPPROCRE.FListeEnter(Sender: TObject);
begin
  BZoomDetail.Enabled := FListe.Cells[0,Fliste.Row]<>'';
end;

procedure TOF_FPRAPPROCRE.BCalculClick(Sender: TObject);
var
  ListeE : TList;
begin
  // Contrôle de la validité de la date de rapprochement
  if ((not IsValidDate(DATERAPPRO.Text)) or
        (StrToDate(DATERAPPRO.Text)>VH^.Encours.Fin) or
        (StrToDate(DATERAPPRO.Text)<VH^.Encours.Deb)) then
    PGIBox ('Date de rapprochement non valide.#10#13Veuillez saisir une date de l''exercice en cours','')
  else
  begin
    ListeE := TList.Create;
    CalculCREECR(ListeE) ;
    AfficheListeEcriture(ListeE) ;

    VideListe(ListeE) ;
    ListeE.Free ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 05/06/2007
Modifié le ... :   /  /    
Description .. : FQ18684
Mots clefs ... : FQ18684 IMPRIMER IMPRESSION
*****************************************************************}
procedure TOF_FPRAPPROCRE.BImprimerClick(Sender: TObject);
begin

  {$IFDEF eAGLCLient} {25.06.07 YMO FQ20840 Impression CWAS}
  TObjEtats.GenereEtatGrille( FListe, Ecran.caption, False) ;
 {$ELSE}
  PrintDBGrid(FListe,PDATERAPPRO,Ecran.caption,'') ;
  {$ENDIF eAGLCLient}

end;

procedure TOF_FPRAPPROCRE.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_F9 then BCALCULClick(nil);
end;

procedure TOF_FPRAPPROCRE.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key=VK_F5) then
  begin;
    FlisteDblClick(nil);
    Key := 0;
  end;
end;



procedure TOF_FPRAPPROCRE.OnDelete;
begin
  inherited;
end;


procedure TOF_FPRAPPROCRE.OnDisplay;
begin
  inherited;
end;

procedure TOF_FPRAPPROCRE.OnUpdate;
begin
  inherited;
end;

procedure TOF_FPRAPPROCRE.OnCancel;
begin
  inherited;
end;

procedure TOF_FPRAPPROCRE.OnLoad;
begin
  inherited;
end;

procedure TOF_FPRAPPROCRE.OnNew;
begin
  inherited;
end;

initialization
  registerclasses([TOF_FPRAPPROCRE]);
end.




