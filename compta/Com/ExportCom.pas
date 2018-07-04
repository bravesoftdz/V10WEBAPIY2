{***********UNITE*************************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 10/10/2001
Modifié le ... :   /  /
Description .. : Export du fichier format PGI S5 format Standard ou etendu
Mots clefs ... :
*****************************************************************}

unit ExportCom;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HStatus, HEnt1, HDB, ComCtrls, Paramsoc, RappType, FE_MAIN,
  StdCtrls, Buttons, ExtCtrls, Hctrls, Mask, hmsgbox, DB, DBTables, UTOB,
  Ent1, UtilTrans, UTOF, HTB97, CALCOLE, uToz, Recordcom,ed_tools,ULibEcriture,
  PGIVersSisco,lookup,RecordComsis,impFicU,ImprimeMaquette, Printers, Mailol;

type
  TOF_ExportCom = Class (TOF)
  procedure OnLoad ; override ;
  procedure OnArgument(stArgument: String); override ;
  procedure bClickValide(Sender: TObject);
  procedure OnClickjournal(Sender : TObject);
  Procedure ExerciceOnChange(Sender :TObject) ;
  procedure DoExoToDateOnChange(Exo: THValComboBox; Date1, Date2: THEdit);
  procedure ChargementCpte(ListeCpteGen : TList;Where : string);
  Procedure AlimJalB(LJB : TList; Where : string) ;
  Procedure ExportComSx;
  procedure ChargementMouvement(var ListeMouvt : TList;Where,FormatEnvoie : string;Var F : TextFile; Pana : string; ListeCpteGen : TList; ListeCpteaux : TList);
  FUNCTION RechRib(cpteaux  : string; var TCompteAux  : PListeAux; var ListeCpteaux,ListeRib : TList) : Boolean;
  procedure RechContact(cpteaux  : string; var TCompteAux  : PListeAux);
  procedure ChargementCpteaux(ListeCpteaux,ListeRib : TList;Where : string);
  procedure ChargementTableLibre(ListeLibre : TList;Where : string);
  procedure ChargementTableSection(ListeSection : TList;Where : string);
  procedure Chargementregle(Listeregle : TList;Where : string);
  procedure ChargementMouvementAnalytique(var ListeMouvt : TList;Where,FormatEnvoie : string; Tdev : TList;Var F : TextFile);
  procedure TraiteListeJournal (Naturejal : string);
  procedure ChargementBalance(ListeMouvt : TList; Var F : TextFile);
  procedure RemplirEcritureBal (var TMvt : PListeMouvt; Compte, comptecol, nature, libelle : string; var MttSolde : double);
  procedure FileClickzip;
  Procedure OnChangeNature(Sender :TObject) ;
  procedure Traitement_Ventil (var TMvt : PListeMouvt; var ListeMouvt : TList; Exo,Compte, Collectif, nature, libelle, Jrl : string);
  procedure OnClickparam(Sender : TObject);
  Procedure OnChangeVers(Sender :TObject) ;
  Procedure ExerciceOnexitDateecr2(Sender :TObject) ;
  Procedure SauvegardeQueParametre(var F: TextFile);
  Procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
  Procedure RendVisibleChamp;
  procedure OnElipsisClickColl(Sender: TObject);
  function  Creationparamcol : Boolean;
  procedure OnClicksanscol(Sender : TObject);
  function  Existedejacoll (Tg : THGRID; compte : string; var Rr : integer) : Boolean;
  procedure RenWhereEcr( var Where: string; Ext : string);
// ajout me 03-05-2002
  Procedure OnChangeChoixCollectif(Sender :TObject) ;
  Procedure OnExitCollectif(Sender :TObject) ;
  procedure bClickAjoutAux(Sender: TObject);
  function  Existedejaaux (Col,T1,T2 : string) : Boolean;
  procedure OnCellEnterFGrille(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
  procedure bClickDelete(Sender : TObject);
  Procedure Onclickaux1(Sender :TObject) ;
  Procedure Onclickaux2(Sender :TObject) ;
  Procedure OnExitAux1(Sender :TObject) ;
  Procedure OnExitAux2(Sender :TObject) ;
  Procedure OnChangeCaraux(Sender :TObject) ;

  private
  Trans                               : TFTransfertcom;
  CarFou,CarCli                       : string;
  TobCol                              : TOB;
  FGrille                             : THGrid;
  OkExoEncours                        : Boolean;
  ListeCpteGen                        : TList;
  WhereOkExoEncours                   : string;
  FTimer                              : TTimer;
  StArg,StArgsav                      : string;
  TOBTiers                            : TOB;
  EtablissUnique                      : Boolean;
  Email                               : string;
  procedure OnMonTimer(Sender: TObject);
end;

implementation

Uses CpteSAV,UtilPGI ;

procedure TOF_ExportCom.OnMonTimer(Sender: TObject);
begin
  FTimer.Enabled := FALSE;
  bClickValide(Sender);
end;

procedure TOF_ExportCom.OnLoad ;
var
Vales     : TStrings;
Q1        : TQuery;
ATob      : TOB;
NomFichT  : string;
tmp       : string;
begin
  Vales := TStringList.Create;
  Vales.Add('STD') ;
  Vales.Add('ETE') ;
  THValComboBox(GetControl('TYPEFORMAT')).Items.Add('STANDARD');
  THValComboBox(GetControl('TYPEFORMAT')).Items.Add('ETENDU');
  THValComboBox(GetControl('TYPEFORMAT')).Values := Vales;
  Vales.Free;
  THValComboBox(GetControl('TYPEFORMAT')).ItemIndex := 0;
  SetControlText ('DATEECR1', '01/01/1900');
  SetControlText ('DATEECR2', '01/01/2099');
  TraiteListeJournal ('');
  SetControlVisible('PEXPORT', FALSE);
  SetControlVisible ('IMAGES1', FALSE);
  SetControlVisible ('GRPBAL', FALSE);
  SetControlVisible ('GRTYPECR', FALSE);
  SetControlVisible ('BNORMAL', FALSE);
  SetControlVisible ('BSIMULE', FALSE);
  SetControlVisible ('BSITUATION', FALSE);
  SetControlVisible('PARAMSISCOII', FALSE);
  THRadioGroup(GetControl('TRANSFERTVERS')).Value := GetParamSocSecur ('SO_CPLIENGAMME', False);
  SetControlText('DATEECR', FormatDateTime('dd/mm/yyyy', now));
// ajout me pour version 590
  SetControlVisible ('HISTOSYNCHRO', FALSE);
  if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1') then
  begin
    THVALComboBox(GetControl('TYPEFORMAT')).Value := 'ETE';
    THValComboBox(GetControl('NATURETRANSFERT')).Items.clear;
    THValComboBox(GetControl('NATURETRANSFERT')).Values.clear;
    THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Dossier');
    THValComboBox(GetControl('NATURETRANSFERT')).Values.add('DOS');
    THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Journal');
    THValComboBox(GetControl('NATURETRANSFERT')).Values.add('JRL');
    THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Synchronisation S1');
    THValComboBox(GetControl('NATURETRANSFERT')).Values.add('SYN');
    THValComboBox(GetControl('NATURETRANSFERT')).ItemIndex := 0;
    RendVisibleChamp;
  end
  else
  if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI') then
    THVALComboBox(GetControl('NATURETRANSFERT')).Value := 'JRL'
  else
    THVALComboBox(GetControl('NATURETRANSFERT')).Value := 'DOS';

  SetControlVisible('PPARAM',  TCheckBox (GetControl ('PARAMDOS')).checked);
  SetControlVisible('PPARAMGENE', not TCheckBox (GetControl ('PARAMDOS')).checked);

//  THGrid(GetControl('LISTESTD')).RowCount := QCount(QPlanRef) + 1;
  CarFou := '0'; CarCli := '9';
  TobCol := TOB.Create('', nil, -1);
  Q1 := OpenSql ('SELECT * from CORRESP Where CR_TYPE="SIS"', TRUE);
  if Q1.EOF then
  begin
            THVALComboBox(GetControl('COLLECTIFS')).Value := 'F';
            THValComboBox(GetControl('CARAUX')).value := '0';
  end
  else
  begin
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '40' then
                THVALComboBox(GetControl('COLLECTIFS')).Value := 'F';
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '41' then
                THVALComboBox(GetControl('COLLECTIFS')).Value := 'C';
           SetControlText ('COLLECT', AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0'));
           THValComboBox(GetControl('CARAUX')).value := Copy (Q1.FindField ('CR_LIBELLE').asstring,1,1);
           SetControlText ('AUX1', AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0'));
           SetControlText ('AUX2', AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0'));
  end;
  while not Q1.EOF do
  begin
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '40' then
                CarFou := copy (Q1.FindField ('CR_LIBELLE').asstring,0,1);
           if Copy (Q1.FindField ('CR_CORRESP').asstring, 0,2) = '41' then
                CarCli := copy (Q1.FindField ('CR_LIBELLE').asstring,0,1);
           ATob := TOB.Create('CORRESP',TobCol,-1);
           ATob.PutValue('CR_CORRESP', AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0'));
           ATob.PutValue('CR_LIBELLE', AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0'));
           ATob.PutValue('CR_ABREGE', AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0'));
           Q1.next;
  end;
  ferme (Q1);
  FGrille := THGrid(GetControl('THGRID'));
  TobCol.PutGridDetail(THGrid(GetControl('THGRID')),False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
  FGrille.OnCellEnter := OnCellEnterFGrille;
// penser à modifier la fiche après
  TToolbarButton97(GetControl('AJOUTAUX')).Flat := TRUE;
  TToolbarButton97(GetControl('BDelete1')).Flat := TRUE;
  THEDIT(GetControl('COLLECT')).EditMask := '';
  THEDIT(GetControl('COLLECT')).Text := '40';
  THEDIT(GetControl('COLLECT')).maxlength := VH^.Cpta[fbGene].Lg;
  THEDIT(GetControl('AUX1')).EditMask := '';
  THEDIT(GetControl('AUX1')).Text := '0';
  THEDIT(GetControl('AUX1')).maxlength := VH^.Cpta[fbGene].Lg;
  THEDIT(GetControl('AUX2')).EditMask := '';
  THEDIT(GetControl('AUX2')).Text := '0';
  THEDIT(GetControl('AUX2')).maxlength := VH^.Cpta[fbGene].Lg;
  TOBTiers := Nil ;
  if stArg <> '' then  // par ligne de commande
  begin
              FTimer.Enabled := TRUE;
              tmp := ReadTokenPipe(stArg, ';');
              tmp := ReadTokenPipe(stArg, ';');
              tmp := ReadTokenPipe(stArg, ';');
              Email := ReadTokenPipe(stArg, ';');
              NomFichT := ReadTokenPipe(stArg, ';');
              if NomFichT <> '' then
              begin
                   if TOBTiers <> nil then TOBTiers.free;
                   TOBTiers := Nil ;
                   TOBTiers :=Tob.create('',Nil,-1) ;
                   if not TobLoadFromFile(NomFichT,nil,TOBTiers) then
                   begin
                        TOBTiers := Nil ;
                        TOBTiers.free;
                   end;
              end;
              tmp := ReadTokenPipe(stArg, ';');
              // pour mettre les ecritures sur le même établissement
              EtablissUnique := (tmp = '-');
              THRadioGroup(GetControl('TRANSFERTVERS')).Value := ReadTokenPipe(stArg, ';');
              THVALComboBox(GetControl('NATURETRANSFERT')).Value := ReadTokenPipe(stArg, ';');   // nom du fichier

              THVALComboBox(GetControl('TYPEFORMAT')).Value := 'ETE';
              SetControlText('FICHENAME',  ReadTokenPipe(stArg, ';'));

              if (THVALComboBox(GetControl('NATURETRANSFERT')).Value = 'JRL')
              or (THVALComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL') then
              begin
                 THVALComboBox(GetControl('TYPEFORMAT')).Value := 'STD';
                 THVALComboBox(GetControl('EXERCICE')).value:= ReadTokenPipe(stArg, ';');
                 SetControlText ('DATEECR1', ReadTokenPipe(stArg, ';'));
                 SetControlText ('DATEECR2', ReadTokenPipe(stArg, ';'));
                 if GetControlText('DATEECR1') = '  /  /    ' then
                    SetControlText ('DATEECR1', '01/01/1900');
                 if GetControlText('DATEECR2') = '  /  /    ' then
                    SetControlText ('DATEECR2', '01/01/2099');
                 SetControlText('Journaux', stArg);
              end;
              stArg := stArgsav;
  end;

end;

procedure TOF_ExportCom.OnElipsisClickColl(Sender: TObject);
begin
if (Copy (GetControlText ('COLLECT'),1,2)) = '40' then
    LookupList(TControl(Sender),'','','G_GENERAL','G_LIBELLE','','',True,-1,'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX WHERE G_GENERAL LIKE "40%" and G_NATUREGENE="COF"');
if (Copy (GetControlText ('COLLECT'),1,2)) = '41' then
    LookupList(TControl(Sender),'','','G_GENERAL','G_LIBELLE','','',True,-1,'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX WHERE G_GENERAL LIKE "41%" and G_NATUREGENE="COC"');
end;

procedure TOF_ExportCom.OnArgument(stArgument: String);
begin
     TToolbarButton97(GetControl('BValider')).Onclick := bClickValide;
     THValComboBox(GetControl('NATUREJRL')).Onclick :=  OnClickjournal;
//    THValComboBox(GetControl('BTWINZIP')).Onclick := BInsertFileClickzip;
     THValComboBox(GetControl('EXERCICE')).OnChange := ExerciceOnChange;
     THRadioGroup(GetControl('TRANSFERTVERS')).Onclick := OnChangeVers;
     THValComboBox(GetControl('NATURETRANSFERT')).OnChange := OnChangeNature;
     TCheckBox (GetControl ('PARAMDOS')).Onclick := OnClickparam;
     THEdit(GetControl('DATEECR2')).Onexit := ExerciceOnexitDateecr2;
     TCheckBox (GetControl ('SANCOLL')).Onclick := OnClicksanscol;
     Ecran.OnKeyDown := OnKeyDownEcran;
     // ajout me 03-05-2002
     THValComboBox(GetControl('COLLECTIFS')).Onchange := OnChangeChoixCollectif;
     THValComboBox(GetControl('CARAUX')).Onchange := OnChangeCaraux;

     THEDIT(GetControl('AUX1')).OnClick := Onclickaux1;
     THEDIT(GetControl('AUX2')).OnClick := Onclickaux2;
     THEDIT(GetControl('AUX1')).OnExit := OnExitAux1;
     THEDIT(GetControl('AUX2')).OnExit := OnExitAux2;

     THEDIT(GetControl('COLLECT')).OnExit := OnExitCollectif;
     TToolbarButton97(GetControl('AJOUTAUX')).Onclick := bClickAjoutAux;
     TToolbarButton97(GetControl('BDelete1')).Onclick := bClickDelete;
     THEDIT(GetControl('COLLECT')).OnElipsisClick := OnElipsisClickColl;
     FTimer := TTimer.Create(nil);
     FTimer.Enabled := FALSE;
     FTimer.Interval:= 1000;
     FTimer.Ontimer := OnMonTimer;
     stArgsav := stArgument;
     stArg := stArgument;

end;

procedure TOF_ExportCom.TraiteListeJournal (Naturejal : string);
var
  Q : TQuery;
  sLejournal,sRequete,sWhere,sFrom, LibAbreg : String;
begin
  THMultiValComboBox(GetControl('JOURNAUX')).Items.clear;
  THMultiValComboBox(GetControl('JOURNAUX')).Values.clear;
  sRequete := 'select J_JOURNAL,J_LIBELLE,J_NATUREJAL ';
  sFrom := 'from JOURNAL ';
  if Naturejal <> '' then
  sWhere := 'where J_NATUREJAL="'+ Naturejal +'" ORDER BY J_JOURNAL';
  Q := OpenSQL(sRequete+sFrom+sWhere,true);
  THMultiValComboBox(GetControl('JOURNAUX')).Items.Add('<<Tous>>');
  THMultiValComboBox(GetControl('JOURNAUX')).Values.add(' ');
  while not Q.EOF do
    begin
    LibAbreg := Q.FindField('J_LIBELLE').AsString;
    sLejournal := Q.FindField('J_JOURNAL').AsString;
    THMultiValComboBox(GetControl('JOURNAUX')).Items.Add(LibAbreg);
    THMultiValComboBox(GetControl('JOURNAUX')).Values.add(sLejournal);
    Q.Next;
    end;
  Ferme(Q);
end;

procedure TOF_ExportCom.OnClickjournal(Sender : TObject);
begin
     TraiteListeJournal (THVALComboBox(GetControl('NATUREJRL')).value);
end;

procedure TOF_ExportCom.OnClicksanscol(Sender : TObject);
begin
  if ((THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI')) then
  begin
       if TCheckBox (GetControl ('SANCOLL')).checked then
       begin
        FGrille.VidePile(FALSE);
        SetControlEnabled ('AJOUTAUX',FALSE);
        SetControlEnabled ('Bdelete1',FALSE);
        SetControlEnabled ('COLLECTIFS',FALSE);
        SetControlEnabled ('COLLECT',FALSE);
        SetControlEnabled ('CARAUX',FALSE);
        SetControlEnabled ('AUX1',FALSE);
        SetControlEnabled ('AUX2',FALSE);
       end
       else
       begin
        if TobCol <> nil then
           TobCol.PutGridDetail(FGrille,False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
        SetControlEnabled ('AJOUTAUX',TRUE);
        SetControlEnabled ('Bdelete1',TRUE);
        SetControlEnabled ('COLLECTIFS',TRUE);
        SetControlEnabled ('COLLECT',TRUE);
        SetControlEnabled ('CARAUX',TRUE);
        SetControlEnabled ('AUX1',TRUE);
        SetControlEnabled ('AUX2',TRUE);
       end;

  end;
end;

procedure TOF_ExportCom.OnClickparam(Sender : TObject);
begin
     SetControlVisible('PPARAM',  TCheckBox (GetControl ('PARAMDOS')).checked);
     SetControlVisible('PPARAMGENE', not TCheckBox (GetControl ('PARAMDOS')).checked);
     if not TCheckBox (GetControl ('PARAMDOS')).checked then
     begin
          SetControlText('EXERCICE', '');
          SetControlText ('DATEECR1', '01/01/1900');
          SetControlText ('DATEECR2', '01/01/2099');
          SetControlText('JOURNAUX', '');
     end;
end;

Procedure TOF_ExportCom.ExerciceOnChange(Sender :TObject) ;
BEGIN
     DoExoToDateOnChange(THVALComboBox(GetControl('EXERCICE')),
     THEdit(GetControl('DATEECR1')),THEdit(GetControl('DATEECR2'))) ;
     if (THVALComboBox(GetControl('EXERCICE')).Text = '<<Tous>>') then
     begin
          SetControlText ('DATEECR1', '01/01/1900');
          SetControlText ('DATEECR2', '01/01/2099');
     end;
END ;


Procedure TOF_ExportCom.OnChangeChoixCollectif(Sender :TObject) ;
var
Vales                  : TStrings;
begin
     if (THVALComboBox(GetControl('COLLECTIFS')).Value = 'F') then
     begin
          THEDIT(GetControl('AUX1')).Text := '0';
          THEDIT(GetControl('AUX2')).Text := '0';
          THEDIT(GetControl('COLLECT')).Text := '40';
          THValComboBox(GetControl('CARAUX')).Items.clear;
          THValComboBox(GetControl('CARAUX')).Values.clear;
          Vales := TStringList.Create;
          Vales.Add('0') ;
          Vales.Add('F') ;
          THValComboBox(GetControl('CARAUX')).Items.Add('0');
          THValComboBox(GetControl('CARAUX')).Items.Add('F');
          THValComboBox(GetControl('CARAUX')).Values := Vales;
          Vales.Free;
     end;
     if (THVALComboBox(GetControl('COLLECTIFS')).Value = 'C') then
     begin
          THEDIT(GetControl('AUX1')).Text := '9';
          THEDIT(GetControl('AUX2')).Text := '9';
          THEDIT(GetControl('COLLECT')).Text := '41';
          THValComboBox(GetControl('CARAUX')).Items.clear;
          THValComboBox(GetControl('CARAUX')).Values.clear;
          Vales := TStringList.Create;
          Vales.Add('9') ;
          Vales.Add('C') ;
          THValComboBox(GetControl('CARAUX')).Items.Add('9');
          THValComboBox(GetControl('CARAUX')).Items.Add('C');
          THValComboBox(GetControl('CARAUX')).Values := Vales;
          Vales.Free;
     end;
end;

Procedure TOF_ExportCom.OnExitCollectif(Sender :TObject) ;
var
Vales     : TStrings;
begin
     SetControlText ('COLLECT', AGauche(GetControlText ('COLLECT'), VH^.Cpta[fbGene].Lg,'0'));

     if (THVALComboBox(GetControl('COLLECTIFS')).Value = 'F')
     and (Copy(GetControlText ('COLLECT'),1,2) <> '40') then
     begin
          PgiInfo ('Le compte collectif doit commencer par 40','Export');
          THEDIT(GetControl('AUX1')).Text := '0';
          THEDIT(GetControl('AUX2')).Text := '0';
          THEDIT(GetControl('COLLECT')).Text := '40';
          THValComboBox(GetControl('CARAUX')).Items.clear;
          THValComboBox(GetControl('CARAUX')).Values.clear;
          Vales := TStringList.Create;
          Vales.Add('0') ;
          Vales.Add('F') ;
          THValComboBox(GetControl('CARAUX')).Items.Add('0');
          THValComboBox(GetControl('CARAUX')).Items.Add('F');
          THValComboBox(GetControl('CARAUX')).Values := Vales;
          Vales.Free; exit;
     end;
     if (THVALComboBox(GetControl('COLLECTIFS')).Value = 'C')
     and (Copy(GetControlText ('COLLECT'),1,2) <> '41') then
     begin
          PgiInfo ('Le compte collectif doit commencer par 41','Export');
          THEDIT(GetControl('AUX1')).Text := '9';
          THEDIT(GetControl('AUX2')).Text := '9';
          THEDIT(GetControl('COLLECT')).Text := '41';
          THValComboBox(GetControl('CARAUX')).Items.clear;
          THValComboBox(GetControl('CARAUX')).Values.clear;
          Vales := TStringList.Create;
          Vales.Add('9') ;
          Vales.Add('C') ;
          THValComboBox(GetControl('CARAUX')).Items.Add('9');
          THValComboBox(GetControl('CARAUX')).Items.Add('C');
          THValComboBox(GetControl('CARAUX')).Values := Vales;
          Vales.Free; exit;
     end;
      if (Copy (GetControlText ('COLLECT'),1,2)) = '40' then
      begin
           if not ExisteSQL ('SELECT G_GENERAL from GENERAUX Where G_GENERAL="'+
           GetControlText ('COLLECT')+'" and G_NATUREGENE="COF"') then
           begin
               PgiInfo ('Le compte collectif n''existe pas dans le plan comptable','Export');
               SetFocusControl('COLLECT') ;
           end;
      end;
      if (Copy (GetControlText ('COLLECT'),1,2)) = '41' then
      begin
           if not ExisteSQL ('SELECT G_GENERAL from GENERAUX Where G_GENERAL="'+
           GetControlText ('COLLECT')+'" and G_NATUREGENE="COC"') then
           begin
               PgiInfo ('Le compte collectif n''existe pas dans le plan comptable','Export');
               SetFocusControl('COLLECT') ;
           end;
      end;
end;

Procedure TOF_ExportCom.OnChangeCaraux(Sender :TObject) ;
var
caraux : string;
begin
         caraux := THValComboBox(GetControl('CARAUX')).value;
         SetControlText ('Aux1',Caraux);
         SetControlText ('Aux2',Caraux);
end;

Procedure TOF_ExportCom.Onclickaux1(Sender :TObject) ;
var
caraux : string;
begin
         caraux := THValComboBox(GetControl('CARAUX')).value;
         SetControlText ('Aux1',Caraux);
end;

Procedure TOF_ExportCom.OnExitAux1(Sender :TObject) ;
var
caraux,car         : string;
begin
         caraux := THValComboBox(GetControl('CARAUX')).value;
         car := Copy(GetControlText ('AUX1'),1,1);
         if (caraux <> car) and (car <> '') then
         begin
              SetControlText ('AUX1', caraux+Copy(GetControlText ('AUX1'),2,9));
              SetControlText ('AUX2', caraux+Copy(GetControlText ('AUX2'),2,9));
         end;
         SetControlText ('AUX1', AGauche(GetControlText ('AUX1'), VH^.Cpta[fbGene].Lg,'0'));
end;

Procedure TOF_ExportCom.Onclickaux2(Sender :TObject) ;
var
caraux : string;
begin
         caraux := THValComboBox(GetControl('CARAUX')).value;
         SetControlText ('Aux2',Caraux);
end;

Procedure TOF_ExportCom.OnExitAux2(Sender :TObject) ;
var
caraux,car         : string;
begin
         caraux := THValComboBox(GetControl('CARAUX')).value;
         car := Copy(GetControlText ('AUX2'),1,1);
         if (caraux <> car) and (car <> '') then
            SetControlText ('AUX2', caraux+Copy(GetControlText ('AUX2'),2,9));
         SetControlText ('AUX2', AGauche(GetControlText ('AUX2'), VH^.Cpta[fbGene].Lg,'0'));
end;

procedure TOF_ExportCom.OnCellEnterFGrille(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var
  CollectifEncours              : string;
  AuxEncours1,AuxEncours2       : string;
  Vales                         : TStrings;
begin
     THValComboBox(GetControl('COLLECTIFS')).Onchange := nil;
     THValComboBox(GetControl('CARAUX')).Onchange := nil;
     CollectifEncours := FGrille.Cells[0, FGrille.Row];
     AuxEncours1 := AGauche(FGrille.Cells[1, FGrille.Row], 10,'0');
     AuxEncours2 := AGauche(FGrille.Cells[2, FGrille.Row], 10,'0');
     if (CollectifEncours <> '') then
     begin
          SetControlText ('COLLECT',AGauche(CollectifEncours, VH^.Cpta[fbGene].Lg,'0'));
          SetControlText('AUX1', Copy(AuxEncours1,1,VH^.Cpta[fbGene].Lg));
          SetControlText('AUX2', Copy(AuxEncours2,1,VH^.Cpta[fbGene].Lg));
          if (Copy(CollectifEncours,1,2) = '41') then
             THVALComboBox(GetControl('COLLECTIFS')).Value := 'C'
          else
              THVALComboBox(GetControl('COLLECTIFS')).Value := 'F';
          THValComboBox(GetControl('CARAUX')).Items.clear;
          THValComboBox(GetControl('CARAUX')).Values.clear;
          Vales := TStringList.Create;
          if copy (CollectifEncours,1,2) = '40' then
          begin
               Vales.Add('0') ; Vales.Add('F') ;
               THValComboBox(GetControl('CARAUX')).Items.Add('0');
               THValComboBox(GetControl('CARAUX')).Items.Add('F');
          end;
          if copy (CollectifEncours,1,2) = '41' then
          begin
               Vales.Add('9') ; Vales.Add('C') ;
               THValComboBox(GetControl('CARAUX')).Items.Add('9');
               THValComboBox(GetControl('CARAUX')).Items.Add('C');
          end;
          THValComboBox(GetControl('CARAUX')).Values := Vales;
          Vales.Free;

          THVALComboBox(GetControl('CARAUX')).Value := Copy(AuxEncours1,1,1);
     end;
     THValComboBox(GetControl('COLLECTIFS')).Onchange := OnChangeChoixCollectif;
     THValComboBox(GetControl('CARAUX')).Onchange := OnChangeCaraux;

end;


procedure TOF_ExportCom.bClickDelete(Sender : TObject);
var
Compte  : string;
iRef    : integer;
OkSup   : Boolean;
begin
  iRef := FGrille.Row;
  if FGrille.Cells[0, iRef] = '' then exit;
  Compte := AGauche(FGrille.Cells[0, iRef], 10,'0');
  if stArg <> '' then OkSup := TRUE
  else
  OkSup  := (PGIAsk ('Confirmez-vous la suppression du paramétrage sélectionné', 'Paramétrage Sisco II ')=mrYes);

  if not OkSup then  exit;

  FGrille.GotoLeBookmark(iRef);
  TOBCol.Detail[iRef-1].Free;
{
  for ii:=TobCol.Detail.Count-1 downto 0 do
  BEGIN
          ATob:=TOBCol.Detail[ii] ;
          if ATob.GetValue ('CR_CORRESP') = Compte then
          begin
               TOBCol.Detail[ii].Free;
          end ;
  END ;
}
  if TobCol <> nil then
  begin
      FGrille.VidePile(FALSE);
      TobCol.PutGridDetail(THGrid(GetControl('THGRID')),False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
  end;

end;


procedure TOF_ExportCom.bClickAjoutAux(Sender: TObject);
var
ATob                  : Tob;
Collectif, Aux1, Aux2 : string;
iRef                  : integer;
begin
   if  TobCol = nil then TobCol := TOB.Create('', nil, -1);

   Collectif := AGauche(GetControlText ('COLLECT'),10,'0');
   Aux1 := AGauche(GetControlText ('AUX1'),10,'0');
   Aux2 := AGauche(GetControlText ('AUX2'),10,'0');
   if (Copy (Collectif,1,2) = '40') and (Copy (Aux1,1,1) <> '0')
   and (Copy (Aux1,1,1) <> 'F') then
   begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;
   if (Copy (Collectif,1,2) = '41') and (Copy (Aux1,1,1) <> '9')
   and (Copy (Aux1,1,1) <> 'C') then
   begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;

   if (Copy (Collectif,1,2) = '40') and (Copy (Aux2,1,1) <> '0')
   and (Copy (Aux2,1,1) <> 'F') then
   begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;
   if (Copy (Collectif,1,2) = '41') and (Copy (Aux2,1,1) <> '9')
   and (Copy (Aux2,1,1) <> 'C') then
   begin PGIInfo ('Tranche de compte incorrecte', ' Paramétrage Sisco II'); exit; end;

   iRef := FGrille.Row;
   if Existedejacoll (THGrid(GetControl('THGRID')), Collectif, iRef) then
   begin
    if (stArg <> '' ) or((stArg = '') and (PGIAsk('Voulez-vous enregistrer les modifications ?','Paramétrage Sisco II ')=mrYes)) then
    begin
         ATob := TOBCol.Detail[iRef-1];
         ATob.PutValue('CR_LIBELLE', AGauche(Aux1, 10,'0'));
         ATob.PutValue('CR_ABREGE', AGauche(Aux2, 10,'0'));
         TobCol.PutGridDetail(THGrid(GetControl('THGRID')),False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');
         exit;
    end
    else exit;
   end;
   if Existedejaaux (Collectif, Aux1, Aux2) then exit;
   Aux1 := GetControlText ('AUX1');
   Aux2 := GetControlText ('AUX2');
   ATob := TOB.Create('CORRESP',TobCol,-1);
   ATob.PutValue('CR_CORRESP', AGauche(Collectif, 10,'0'));
   ATob.PutValue('CR_LIBELLE', AGauche(Aux1, 10,'0'));
   ATob.PutValue('CR_ABREGE', AGauche(Aux2, 10,'0'));
   TobCol.PutGridDetail(THGrid(GetControl('THGRID')),False,False,'CR_CORRESP;CR_LIBELLE;CR_ABREGE');

end;

Procedure TOF_ExportCom.ExerciceOnexitDateecr2(Sender :TObject) ;
Var D1,D2 : TDateTime ;
    Q     : TQuery;
    Okok  : boolean ;
    EXO   : String3;
BEGIN
   EXO := THVALComboBox(GetControl('EXERCICE')).value;
   if (THVALComboBox(GetControl('EXERCICE')).Text <> '<<Tous>>')
        and (THVALComboBox(GetControl('EXERCICE')).Text <> '') then
    begin
    Okok:=True ; D1:=Date ; D2:=Date ;
    If EXO=VH^.Precedent.Code Then BEGIN D1:=VH^.Precedent.Deb ; D2:=VH^.Precedent.Fin ; END Else
    If EXO=VH^.EnCours.Code Then BEGIN D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ; END Else
    If EXO=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END Else
    BEGIN
    Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
    if Not Q.EOF then
      BEGIN
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      END else Okok:=False ;
      Ferme(Q) ;
    END;
    if Okok then
    begin
         if (D1 <= strTodate(THEdit(GetControl('DATEECR1')).text)) and
         (D2  >= strTodate(THEdit(GetControl('DATEECR2')).text)) then
         begin
              if (strTodate(THEdit(GetControl('DATEECR1')).text)) <=
                 (strTodate(THEdit(GetControl('DATEECR2')).text)) then exit
              else
                  DoExoToDateOnChange(THVALComboBox(GetControl('EXERCICE')),
                  THEdit(GetControl('DATEECR1')),THEdit(GetControl('DATEECR2'))) ;
         end
         else
           DoExoToDateOnChange(THVALComboBox(GetControl('EXERCICE')),
           THEdit(GetControl('DATEECR1')),THEdit(GetControl('DATEECR2'))) ;
    end;
    end;
END ;

procedure TOF_ExportCom.RendVisibleChamp;
var
Typef     : Boolean;
Q1        : TQuery;
begin
  typef := ((THRadioGroup(GetControl('TRANSFERTVERS')).Value <> 'SI')) and
     ((THRadioGroup(GetControl('TRANSFERTVERS')).Value <> 'S1')) and
     (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL');

  if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1') then
   THVALComboBox(GetControl('TYPEFORMAT')).Value := 'ETE'
  else
      if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI') then
           THVALComboBox(GetControl('TYPEFORMAT')).Value := 'STD'
  else
  if THValComboBox(GetControl('NATURETRANSFERT')).Value = 'DOS' then
    THVALComboBox(GetControl('TYPEFORMAT')).Value := 'ETE';
  SetControlEnabled ('TYPEFORMAT', typef);
  SetControlEnabled ('TTYPEFORMAT', typef);
  SetControlVisible ('PARAMDOS', typef);
  if ((THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI')) or
     ((THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1')) then
  TCheckBox (GetControl ('PARAMDOS')).checked := typef;
  SetControlEnabled ('TYPEFORMAT', not (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL'));
  SetControlEnabled ('TTYPEFORMAT', not (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL'));
  if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'DOS')
  or (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
  begin
       GrisechampCombo (THValComboBox(GetControl('EXERCICE')), FALSE);
       SetControlEnabled ('TEXERCICE', FALSE);
       GrisechampCombo (THValComboBox(GetControl('NATUREJRL')), FALSE);
       SetControlEnabled ('TNATUREJRL', FALSE);
       GrisechampCombo (THValComboBox(GetControl('JOURNAUX')), FALSE);
       SetControlEnabled ('TJOURNAUX', FALSE);
       GrisechampCombo (THValComboBox(GetControl('FETAB')), FALSE);
       SetControlEnabled ('TFETAB', FALSE);
       GrisechampEdit (THEdit(GetControl('DATEECR1')), FALSE);
       SetControlEnabled ('TDATEECR1', FALSE);
       GrisechampEdit (THEdit(GetControl('DATEECR2')), FALSE);
       SetControlEnabled ('TDATEECR2', FALSE);
       THVALComboBox(GetControl('TYPEFORMAT')).Value := 'ETE';
  end
  else
  begin
       // ajout me 09-12-2002
       if (THValComboBox(GetControl('NATURETRANSFERT')).Value <> 'EXE') then
       begin
                 THValComboBox(GetControl('EXERCICE')).Items.clear;
                 THValComboBox(GetControl('EXERCICE')).Values.clear;
                 THValComboBox(GetControl('EXERCICE')).Items.Add('<<Tous>>');
                 THValComboBox(GetControl('EXERCICE')).Values.add(' ');

                 Q1 := OpenSql ('SELECT EX_LIBELLE,EX_EXERCICE from EXERCICE ',FALSE);
                 While not Q1.EOF do
                 begin
                         THValComboBox(GetControl('EXERCICE')).Items.Add(Q1.FindField('EX_LIBELLE').asstring);
                         THValComboBox(GetControl('EXERCICE')).Values.add(Q1.FindField('EX_EXERCICE').asstring);
                         Q1.next;
                 end;
                 ferme (Q1);
       end;
       if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL')
       // ajout me 09-12-2002
       or (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'EXE')then
       begin
           GrisechampCombo (THValComboBox(GetControl('EXERCICE')), TRUE);
           SetControlEnabled ('TEXERCICE', TRUE);
           GrisechampCombo (THValComboBox(GetControl('NATUREJRL')), FALSE);
           SetControlEnabled ('TNATUREJRL', FALSE);
           GrisechampCombo (THValComboBox(GetControl('JOURNAUX')), FALSE);
           SetControlEnabled ('TJOURNAUX', FALSE);
           THVALComboBox(GetControl('TYPEFORMAT')).Value := 'STD';
       // ajout me 09-12-2002
           if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'EXE') then
           begin
                 THValComboBox(GetControl('EXERCICE')).Items.clear;
                 THValComboBox(GetControl('EXERCICE')).Values.clear;
                 Q1 := OpenSql ('SELECT EX_LIBELLE,EX_EXERCICE from EXERCICE WHERE EX_ETATCPTA="OUV"  ORDER BY EX_EXERCICE',FALSE);
                 if not Q1.EOF  then
                 begin
                         THValComboBox(GetControl('EXERCICE')).Items.Add(Q1.FindField('EX_LIBELLE').asstring);
                         THValComboBox(GetControl('EXERCICE')).Values.add(Q1.FindField('EX_EXERCICE').asstring);
                 end;
                 ferme (Q1);
                 GrisechampCombo (THValComboBox(GetControl('FETAB')), FALSE);
                 SetControlEnabled ('TFETAB', FALSE);
                 GrisechampEdit (THEdit(GetControl('DATEECR1')), FALSE);
                 SetControlEnabled ('TDATEECR1', FALSE);
                 GrisechampEdit (THEdit(GetControl('DATEECR2')), FALSE);
                 SetControlEnabled ('TDATEECR2', FALSE);
           end
           else
           begin
                 GrisechampCombo (THValComboBox(GetControl('FETAB')), TRUE);
                 SetControlEnabled ('TFETAB', TRUE);
                 GrisechampEdit (THEdit(GetControl('DATEECR1')), TRUE);
                 SetControlEnabled ('TDATEECR1', TRUE);
                 GrisechampEdit (THEdit(GetControl('DATEECR2')), TRUE);
                 SetControlEnabled ('TDATEECR2', TRUE);
           end;
       end
       else
       begin
           GrisechampCombo (THValComboBox(GetControl('EXERCICE')), TRUE);
           SetControlEnabled ('TEXERCICE', TRUE);
           GrisechampCombo (THValComboBox(GetControl('NATUREJRL')), TRUE);
           SetControlEnabled ('TNATUREJRL', TRUE);
           GrisechampCombo (THValComboBox(GetControl('JOURNAUX')), TRUE);
           SetControlEnabled ('TJOURNAUX', TRUE);
           GrisechampCombo (THValComboBox(GetControl('FETAB')), TRUE);
           SetControlEnabled ('TFETAB', TRUE);
           GrisechampEdit (THEdit(GetControl('DATEECR1')), TRUE);
           SetControlEnabled ('TDATEECR1', TRUE);
           GrisechampEdit (THEdit(GetControl('DATEECR2')), TRUE);
           SetControlEnabled ('TDATEECR2', TRUE);
       end;
  end;
  if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI') then
  begin
    SetControlEnabled ('TYPEFORMAT', FALSE);
    SetControlEnabled ('TTYPEFORMAT', FALSE);
    SetControlVisible ('GRPBAL', FALSE);
    SetControlVisible ('GRTYPECR', FALSE);
    SetControlVisible ('PARAMSISCOII', TRUE);
    SetControlVisible ('BDelete1', TRUE);
  end
  else
  begin
    SetControlVisible('PARAMSISCOII', FALSE);
    SetControlVisible ('BDelete1', FALSE);
  end;
  SetControlVisible ('TDATEECR', (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN'));
  SetControlVisible ('DATEECR', (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN'));

end;

Procedure TOF_ExportCom.OnChangeNature(Sender :TObject) ;
var
OkBal,OkType     : Boolean;
begin
  OkBal :=  (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL');
  OkType :=  (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL')
  or (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'JRL');

  SetControlVisible ('GRPBAL', OkBal);
  SetControlVisible ('GRTYPECR', OkType);

  if (THRadioGroup(GetControl('TRANSFERTVERS')).Value <> 'SI') then
  begin
     SetControlVisible ('BSIMULE', OkType);
     SetControlVisible ('BSITUATION', OkType);
     SetControlVisible ('BNORMAL', OkType);
  end;
  RendVisibleChamp;

  if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
  begin
   SetControlVisible ('HISTOSYNCHRO', TRUE);
   THVALComboBox(GetControl('TYPEFORMAT')).Value := 'ETE';
   THRadioGroup(GetControl('TRANSFERTVERS')).Value := 'S1';
   SetControlEnabled ('TYPEFORMAT', FALSE);
   SetControlEnabled ('TTYPEFORMAT', FALSE);
// ajout me pour version 590
    THValComboBox(GetControl('HISTOSYNCHRO')).Items.clear;
    THValComboBox(GetControl('HISTOSYNCHRO')).Values.clear;
    THValComboBox(GetControl('HISTOSYNCHRO')).Items.Add('<<Encours>>');
    THValComboBox(GetControl('HISTOSYNCHRO')).Values.add('X');
    if (GetParamSocSecur ('SO_CPDATESYNCHRO1', False) <=  GetParamSocSecur ('SO_CPDATESYNCHRO2', False)) then
    begin
         if (GetParamSocSecur('SO_CPDATESYNCHRO1', False) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO1', False) <> iDate1900) then
         begin
              THValComboBox(GetControl('HISTOSYNCHRO')).Items.Add(FormatDateTime('dddd dd mmmm yyyy "à" hh:nn', GetParamSocSecur('SO_CPDATESYNCHRO1', False)));
              THValComboBox(GetControl('HISTOSYNCHRO')).Values.add('0');
         end;
         if (GetParamSocSecur('SO_CPDATESYNCHRO2', False) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO2', False) <> iDate1900) then
         begin
              THValComboBox(GetControl('HISTOSYNCHRO')).Items.Add(FormatDateTime('dddd dd mmmm yyyy "à" hh:nn', GetParamSocSecur('SO_CPDATESYNCHRO2', False)));
              THValComboBox(GetControl('HISTOSYNCHRO')).Values.add('1');
         end;
    end
    else
    begin
         if (GetParamSocSecur('SO_CPDATESYNCHRO2', False) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO2', False) <> iDate1900) then
         begin
              THValComboBox(GetControl('HISTOSYNCHRO')).Items.Add(FormatDateTime('dddd dd mmmm yyyy "à" hh:nn', GetParamSocSecur('SO_CPDATESYNCHRO2', False)));
              THValComboBox(GetControl('HISTOSYNCHRO')).Values.add('1');
         end;
         if (GetParamSocSecur('SO_CPDATESYNCHRO1', False) <> 0) and (GetParamSocSecur('SO_CPDATESYNCHRO1', False) <> iDate1900) then
         begin
              THValComboBox(GetControl('HISTOSYNCHRO')).Items.Add(FormatDateTime('dddd dd mmmm yyyy "à" hh:nn', GetParamSocSecur('SO_CPDATESYNCHRO1', False)));
              THValComboBox(GetControl('HISTOSYNCHRO')).Values.add('0');
         end;
    end;
  end
  else
       SetControlVisible ('HISTOSYNCHRO', FALSE);
end;

Procedure TOF_ExportCom.OnChangeVers(Sender :TObject) ;
begin
  SetControlProperty('FICHENAME','DataType','SAVEFILE(PG*.TRA)');
  if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1') then
  begin
    THValComboBox(GetControl('NATURETRANSFERT')).Items.clear;
    THValComboBox(GetControl('NATURETRANSFERT')).Values.clear;
    THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Dossier');
    THValComboBox(GetControl('NATURETRANSFERT')).Values.add('DOS');
    THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Journal');
    THValComboBox(GetControl('NATURETRANSFERT')).Values.add('JRL');
    THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Synchronisation S1');
    THValComboBox(GetControl('NATURETRANSFERT')).Values.add('SYN');
    THValComboBox(GetControl('NATURETRANSFERT')).ItemIndex := 0;
    THValComboBox(GetControl('NATURETRANSFERT')).Value := 'DOS';
  end
  else
  begin
      if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI') then
      begin
        THValComboBox(GetControl('NATURETRANSFERT')).Items.clear;
        THValComboBox(GetControl('NATURETRANSFERT')).Values.clear;
        THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Balance');
        THValComboBox(GetControl('NATURETRANSFERT')).Values.add('BAL');
        THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Journal');
        THValComboBox(GetControl('NATURETRANSFERT')).Values.add('JRL');
        // ajout me 9-12-2002 SISCO EXO
        THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Exercice');
        THValComboBox(GetControl('NATURETRANSFERT')).Values.add('EXE');

        THValComboBox(GetControl('NATURETRANSFERT')).ItemIndex := 1;
        THValComboBox(GetControl('NATURETRANSFERT')).Value := 'JRL';
        SetControlProperty('FICHENAME','DataType','SAVEFILE(SI*.TRT)');
      end
      else
      begin
        THValComboBox(GetControl('NATURETRANSFERT')).Items.clear;
        THValComboBox(GetControl('NATURETRANSFERT')).Values.clear;
        THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Dossier');
        THValComboBox(GetControl('NATURETRANSFERT')).Values.add('DOS');
        THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Balance');
        THValComboBox(GetControl('NATURETRANSFERT')).Values.add('BAL');
        THValComboBox(GetControl('NATURETRANSFERT')).Items.Add('Journal');
        THValComboBox(GetControl('NATURETRANSFERT')).Values.add('JRL');
        THValComboBox(GetControl('NATURETRANSFERT')).ItemIndex := 0;
        THValComboBox(GetControl('NATURETRANSFERT')).Value := 'DOS';

      end;
  end;
  RendVisibleChamp;
end;

procedure TOF_ExportCom.DoExoToDateOnChange(Exo: THValComboBox; Date1, Date2: THEdit);
BEGIN
if (Exo=nil) or (Date1=nil) or (Date2=nil) then exit;
  ExoToDates(Exo.Value, Date1, Date2);
END;


procedure TOF_ExportCom.bClickValide(Sender: TObject);
var
Bufnature    : string;
retour       : Boolean;
FileName     : string;
Filearchive  : string;
Datearr      : TDateTime;
Q1           : TQuery;
FichierImp   : string;
Dates        : string;
ListeEmail   : TStringList;
begin
        OkExoEncours := FALSE;
        Trans.Jr               := '';
        Dates := '';
        Trans.Serie := THRadioGroup(GetControl('TRANSFERTVERS')).Value;
        if (THVALComboBox(GetControl('EXERCICE')).Text <> '<<Tous>>')
        and (THVALComboBox(GetControl('EXERCICE')).Text <> '') then
            Trans.Exo              := THVALComboBox(GetControl('EXERCICE')).value;
        if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI') and (Trans.Exo = '') then
        begin
                 PGIBox ('Veuillez ne sélectionner qu''un seul exercice.','Export');
                 if TobCol <> nil then
                 begin TobCol.free;  TobCol := nil; end;
                 exit;
        end;
        if (Trans.Serie = 'SI') then// sisco
        begin
             if VH^.Cpta[fbGene].Lg <> VH^.Cpta[fbAux].Lg then
             begin
                  PGIBox ('Le transfert est impossible.#10#13Les comptes généraux et auxiliaires sont de longueurs différentes','Export Sisco II');
                  if TobCol <> nil then
                  begin TobCol.free;  TobCol := nil; end;
                  exit;
             end;
             if not Creationparamcol then exit;
             if not ExisteSQL('SELECT * from CORRESP Where CR_TYPE="SIS"') then
             begin
                  if not TCheckBox (GetControl ('SANCOLL')).checked  then
                  begin
                       if GetParamSocSecur ('SO_DEFCOLFOU', False) = '' then
                       begin
                        PGiInfo ('Veuillez renseigner le collectif fournisseur dans les paramètres société.', 'Transfert Sisco II');
                        exit;
                       end;
                       if GetParamSocSecur ('SO_DEFCOLCLI', False) = '' then
                       begin
                        PGiInfo ('Veuillez renseigner le collectif client dans les paramètres sociétés.','Transfert Sisco II');
                        exit;
                       end;
                  end;
             end;
        end;
        if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'SI') and
            (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'DOS') then
        begin
                 PGIBox ('Transfert du dossier vers Sisco II impossible.','Export');
                 THValComboBox(GetControl('NATURETRANSFERT')).Value := 'JRL';
                  if TobCol <> nil then
                  begin TobCol.free;  TobCol := nil; end;
                  exit;
        end;
        if (TCheckBox (GetControl ('PARAMDOS')).checked) and
            (THValComboBox(GetControl('NATURETRANSFERT')).Value <> 'BAL') then
        begin
              PGIBox ('Vous devez être en mode balance pour effectuer l''export des paramètres dossier.','Export');
              THValComboBox(GetControl('NATURETRANSFERT')).Value := 'BAL';
              exit;
        end;

        if (THMultiVALComboBox(GetControl('Journaux')).Text <> '<<Tous>>')
        and (THMultiVALComboBox(GetControl('Journaux')).Text <> '') then
           Trans.Jr            := THMultiVALComboBox(GetControl('Journaux')).Text;
        Trans.FichierSortie    := GetControlText('FICHENAME');
        Trans.Dateecr1         := StrToDate(GetControlText('DATEECR1'));
        Trans.Dateecr2         := StrToDate(GetControlText('DATEECR2'));
        Trans.Etabi            :=  GetControlText('FETAB');
        Trans.balance          := (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL');
        Trans.Pgene            := GetControlText('PGEN');
        Trans.pecr             := GetControlText('PECR');
        Trans.ptiers           := GetControlText('PTIERS');
        Trans.pana             := GetControlText('PANA');
        Trans.psection         := GetControlText('PSECTION');
        Trans.pjournaux        := GetControlText('PJRL');
        Trans.Typearchive      := THValComboBox(GetControl('NATURETRANSFERT')).Value;
        ListeCpteGen := nil;
        if ExisteSQl ('SELECT D_PARITEEURO FROM DEVISE Where D_PARITEEURO <= 0') then
        begin
                PGIInfo ('Problème parité fixe EURO d''une devise,  elle ne doit pas être à zéro et / ou négative.'+#10#13+
                'Veuillez modifier les paramètres comptables' ,'Export');
                exit;
        end;
(*        if Trans.Serie = 'S5' then
        begin
                 if (GetParamSocSecur ('SO_CPLIENGAMME') <> 'S5') then
                 begin
                    PGIBox ('Vous n''êtes pas en liaison avec S5'+#10#13
                    +' Pour effectuer un export de données vers S5,'+#10#13
                    +' veuillez renseigner l''option "liaison avec une comptabilité".'+#10#13
                    +' Accès : "Révision à distance" des paramètres sociétés de la comptabilité,'+#10#13
                    +' l''option Liaison avec une comptabilité','Export');
                      exit;
                 end;
        end;
*)
        if Trans.Serie = 'S3' then
        begin
                 if (GetParamSocSecur ('SO_CPLIENGAMME', False) <> 'S3') then
                 begin
                    PGIBox ('Vous n''êtes pas en liaison avec S3'+#10#13
                    +' Pour effectuer un export de données vers S3,'+#10#13
                    +' veuillez renseigner l''option "liaison avec une comptabilité".'+#10#13
                    +' Accès : "Révision à distance" des paramètres sociétés de la comptabilité,'+#10#13
                    +' l''option Liaison avec une comptabilité','Export');
                      exit;
                 end;
        end;
        if Trans.Serie = 'S7' then
        begin
                 if (GetParamSocSecur ('SO_CPLIENGAMME', False) <> 'S7') then
                 begin
                    PGIBox ('Vous n''êtes pas en liaison avec S7'+#10#13
                    +' Pour effectuer un export de données vers S7,'+#10#13
                    +' veuillez renseigner l''option "liaison avec une comptabilité".'+#10#13
                    +' Accès : "Révision à distance" des paramètres sociétés de la comptabilité,'+#10#13
                    +' l''option Liaison avec une comptabilité','Export');
                      exit;
                 end;
        end;

        if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1')
        or (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
        begin
           if (GetParamSocSecur ('SO_CPLIENGAMME', False) = '')  then
              SetParamsoc ('SO_CPLIENGAMME', 'S1')
           else
           if (GetParamSocSecur ('SO_CPLIENGAMME', False) <> 'S1')  and
           (GetParamSocSecur ('SO_CPLIENGAMME', False) <> 'S2') then
           begin
                PGIBox ('Vous n''êtes pas en liaison avec S1'+#10#13
                +' Pour effectuer un export de données vers S1,'+#10#13
                +' veuillez renseigner l''option "liaison avec une comptabilité".'+#10#13
                +' Accès : "Révision à distance" des paramètres sociétés de la comptabilité,'+#10#13
                +' l''option Liaison avec une comptabilité','Export');
                exit;
           end;
           if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') and
              (GetParamSocSecur ('SO_CPDECLOTUREDETAIL', False)= TRUE) then
           begin
                PGIBox ('L''export de type synchronisation sera impossible car une suppression du journal d''à nouveaux a été effectué dans la comptabilité.'+#10#13+
                ' Veuillez sélectionner l''export type "dossier"','Export');
                exit;
           end;

           if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') and
              (GetParamSocSecur ('SO_CPSYNCHROSX', False)= TRUE) then
           begin

               if (GetParamSocSecur ('SO_CPDATESYNCHRO1', False) <=  GetParamSocSecur ('SO_CPDATESYNCHRO2', False)) then
               begin
                  if (GetParamSocSecur('SO_CPDATESYNCHRO2', False) <> 0) and (GetParamSocSecur ('SO_CPDATESYNCHRO2', False) <> iDate1900) then Dates := 'le '+ DateToStr(GetParamSocSecur ('SO_CPDATESYNCHRO2', False));
               end else   if (GetParamSocSecur('SO_CPDATESYNCHRO1', False) <> 0) and (GetParamSocSecur ('SO_CPDATESYNCHRO1', False) <> iDate1900) then Dates := 'le '+ DateToStr(GetParamSocSecur ('SO_CPDATESYNCHRO1', False));

               retour := (PGIAsk ('Un envoi de type synchronisation a été effectué '+ Dates +'.'+ #10#13
                +' A ce jour aucune réception n''a été enregistrée.'+ #10#13+' Confirmez-vous l''export ?','Synchronisation S1')=mrYes);
                if not retour then exit;
           end;

           ListeCpteGen:=TList.Create;
           ChargementCpte(ListeCpteGen, '');
           if CompteEstalpha (ListeCpteGen) then
           begin
              PgiBox ('Export impossible.'+#10#13+
              ' Vous avez des comptes généraux alpha numériques.', 'Export');
              LibererGen (ListeCpteGen);
              exit;
           end;
           Trans.TypeFormat := 'ETE';
           Q1 := OpenSQL ('SELECT * FROM ECRITURE WHERE E_ECRANOUVEAU="H" ORDER BY E_DATECOMPTABLE DESC',True);
           if not Q1.EOF then
             begin
                    Trans.Exo := Q1.FindField('E_EXERCICE').AsString;
                    (* pas de question pour le moment OkExoEncours  := (PGIAsk ('Voulez-vous exporter uniquement l''exercice encours', 'Export S1')=mrYes);
                    if not OkExoEncours then Trans.Exo := ''
                    *)
                    OkExoEncours := TRUE;
             end;
             ferme (Q1);
             if OkExoEncours then
             begin
                        Q1 := OpenSql ('SELECT * from EXERCICE WHERE EX_EXERCICE > "'+ Trans.Exo+'"',FALSE);
                        While not Q1.EOF do
                        begin
                                  Trans.Exo := Trans.Exo+';'+Q1.FindField ('EX_EXERCICE').asstring;
                                  Q1.next;
                        end;
                        ferme (Q1);
             end;
        end
        else
           Trans.TypeFormat       := THVALComboBox(GetControl('TYPEFORMAT')).Value;
        Trans.naturejal        := THVALComboBox(GetControl('NATUREJRL')).value;

        if(THValComboBox(GetControl('NATURETRANSFERT')).Value = 'DOS') then
        Bufnature := 'Transfert : Dossier'
        else
            if(THValComboBox(GetControl('NATURETRANSFERT')).Value = 'BAL') then
               Bufnature := 'Transfert : Balance '
               else if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'JRL') then
                   Bufnature := 'Transfert : Journal'
               else if(THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
                   Bufnature := 'Transfert : Synchronisation S1';
        if Trans.Serie = 'SI' then
        begin
                 FileName := Trans.FichierSortie;
                 FileName := ReadTokenPipe (Filename, '.');
                 FileName := FileName + '.TRT';
                 SetControlText('FICHENAME', FileName);
                 Trans.FichierSortie :=  FileName;
                 ListeCpteGen:=TList.Create;
                 if (GetParamSocSecur ('SO_CPLIENGAMME', False) = '')  then
                    SetParamsoc ('SO_CPLIENGAMME', 'SI')
                 else
                 if (GetParamSocSecur ('SO_CPLIENGAMME', False) <> 'SI') then
                 begin
                    PGIBox ('Vous n''êtes pas en liaison avec Sisco'+#10#13
                    +' Pour effectuer un export de données vers Sisco,'+#10#13
                    +' veuillez renseigner l''option "liaison avec une comptabilité".'+#10#13
                    +' Accès : "Révision à distance" des paramètres sociétés de la comptabilité,'+#10#13
                    +' l''option Liaison avec une comptabilité','Export');
                      exit;
                 end;

                 ChargementCpte(ListeCpteGen, '');
                 if CompteEstalpha (ListeCpteGen) then
                 begin
                    PgiBox ('Export impossible.'+#10#13+' Vous avez des comptes généraux alpha numériques.', 'Export');
                    LibererGen (ListeCpteGen);
                    exit;
                 end;

        end;

        if TCheckBox (GetControl ('ZIPPE')).checked then
        begin
             FileName := GetControlText('FICHENAME');
             Filearchive := ReadTokenPipe (Filename, '.');
             Filearchive := Filearchive + '.zip';
             FileName :=  Filearchive;
        end
        else
             FileName :=  GetControlText('FICHENAME');

        if stArg <> '' then Retour := TRUE
        else
        retour := (PGIAsk ('Confirmez-vous le traitement ?', 'Export')=mrYes);

        if Retour then
        begin
             Datearr := StrToDate (GetControlText('DATEECR'));
             DateFichierClo :=  FormatDateTime('ddmmyyyy',Datearr);

             if (Trans.Serie = 'SI') then// sisco
             begin
                 if Trans.Dateecr1 = iDate1900 then
                 begin
                      Trans.Dateecr1 := VH^.Encours.Deb;
                      Trans.Dateecr2 := VH^.Encours.Fin;
                 end;
                 // ajout me 15-07-2003 traitement stat
                 Trans.psection         := THValComboBox(GetControl('CP_STAT')).value;
                 SetControlVisible ('IMAGES1', FALSE);
                 SetControlVisible ('IMAGES5', FALSE);
                 SetControlVisible('PEXPORT',true);
                 SetActiveTabSheet('PEXPORT');
                 if TobCol <> nil then
                 begin TobCol.free;  TobCol := nil; end;
                 AfficheListeCom('Debut Export ', TListBox(GetControl('LISTEEXPORT')));
                 if not TransfertPgiVersSisco(Trans.FichierSortie,TListBox(GetControl('LISTEEXPORT')), Trans, TCheckBox (GetControl ('SANCOLL')).checked, CarCli,CarFou) then
                    AfficheListeCom('Export annulé ', TListBox(GetControl('LISTEEXPORT')))
                 else
                 begin
                     AfficheListeCom('Export terminé ', TListBox(GetControl('LISTEEXPORT')));
                     FichierImp := 'ListeCom'+FormatDateTime('yyyymmddhhnn',NowH)+'.txt';
                     EcrireDansfichierListeCom (FichierImp, TListBox(GetControl('LISTEEXPORT')));
                     if stArg = '' then
                     begin
                          if PGIAsk('Export terminé : voulez-vous imprimer le rapport ?','Export')=mrYes then
                             ControlTextToPrinter(FichierImp,poPortrait);
                     end
                     else
                     begin
                           if Email <> '' then
                           begin
                               ListeEmail := TStringList.Create;
                               ListeEmail.add ('Ci-joint le rapport de Comsx');
                               SendMail('Rapport Comsx', Email, '', ListeEmail, FichierImp, true, 1);
                               ListeEmail.free;
                           end;
                     end;
                 end;
             end
             else
               ExportComSx;
             if TCheckBox (GetControl ('ZIPPE')).checked then
               FileClickzip;
        end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : M.ENTRESSANGLE
Créé le ...... : 22/03/2002
Modifié le ... :   /  /
Description .. : Fonction export Comsx
Mots clefs ... : EXPORTCOMSX
*****************************************************************}

Procedure TOF_ExportCom.ExportComSx;
var
ListeJalB, ListeMouvt, ListeCpteaux : TList;
ListeLibre,ListeSection,ListeRib    : TList;
Listeregle              : TList;
FichierIE               : TextFile;
Where                   : string;
Jourx                   : string;
Typeano                 : string;
FichierImp,Rep          : string;
ListeEmail              : TStringList;
begin
    ListeCpteaux := nil;
    if Trans.FichierSortie='' then
      begin
        PgiBox('Veuillez choisir un fichier d''export #10'+Trans.FichierSortie,'Export') ;
        Exit ;
      end ;
    if Not _BlocageMonoPoste(True) then Exit ;

    AssignFile(FichierIE,Trans.FichierSortie) ;
    Rewrite(FichierIE) ;
    if IoResult<>0 then
      begin
        PgiBox('Impossible d''écrire dans le fichier #10'+Trans.FichierSortie,'Export') ;
        Exit ;
      end ;
    try
         SetControlVisible ('IMAGES1', FALSE);
         SetControlVisible ('IMAGES5', FALSE);
         if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1')
         or (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
            SetControlVisible ('IMAGES1', TRUE)
         else
             SetControlVisible ('IMAGES5', TRUE);

         SetControlVisible('PEXPORT',true);
         SetActiveTabSheet('PEXPORT');

       // TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('EXPORT'));
        // entete
        if TCheckBox (GetControl ('PARAMDOS')).checked  then
          Ecritureentete(FichierIE, THValComboBox(GetControl('NATURETRANSFERT')).Value, Trans.TypeFormat,Trans.Serie,'X')
        else
          Ecritureentete(FichierIE, THValComboBox(GetControl('NATURETRANSFERT')).Value, Trans.TypeFormat,Trans.Serie,'-');

          AfficheListeCom('Paramètres généraux', TListBox(GetControl('LISTEEXPORT')));
        // pour les paramètres généraux
          Ecrituregeneraux(FichierIE);

          AfficheListeCom('Exercice', TListBox(GetControl('LISTEEXPORT')));
        // pour les exercices
          Where := '';
          WhereOkExoEncours := '';
          if Trans.Exo <> '' then
             Where := ' Where ' +RendCommandeExo ('EX',Trans.Exo);
          EcritureExercice(FichierIE, Where);
          Where := '';

          // si l'export uniquement des paramètres dossiers
          if TCheckBox (GetControl ('PARAMDOS')).checked  then
          begin
               SauvegardeQueParametre(FichierIE);
               AfficheListeCom('Export terminé ', TListBox(GetControl('LISTEEXPORT')));
               FichierImp := 'ListeCom'+FormatDateTime('yyyymmddhhnn',NowH)+'.txt';
               EcrireDansfichierListeCom (FichierImp, TListBox(GetControl('LISTEEXPORT')));
               if stArg = '' then
               if PGIAsk('Export terminé : voulez-vous imprimer le rapport ?','Export')=mrYes then
                ControlTextToPrinter(FichierImp,poPortrait);

               exit;
          end;

        if (Trans.Serie <> 'S1') then
        begin
          // pour table libre
          AfficheListeCom('Tables Libres', TListBox(GetControl('LISTEEXPORT')));
           ListeLibre := TList.Create;
           ChargementTableLibre(ListeLibre, '');
           Ecrituretablelibre(ListeLibre,FichierIE);
           ListeLibre.Free ;

             // section analytique
           if ((Trans.Pana = 'X') and (Trans.balance <> TRUE))  or ((GetControlText('BANA') = 'X') and (Trans.balance = TRUE)) then
           begin
               AfficheListeCom('Sections analytiques', TListBox(GetControl('LISTEEXPORT')));
               ListeSection := TList.Create;
               ChargementTableSection(ListeSection, '');
               Ecrituretablesection(ListeSection,FichierIE);
               ListeSection.Free ;
           end;
             // établissement
              AfficheListeCom('Etablissements', TListBox(GetControl('LISTEEXPORT')));
              if (EtablissUnique) and (stArg <> '') then
                 EcritureEtablissement(FichierIE, ' Where ET_ETABLISSEMENT="'+GetParamSocSecur ('SO_ETABLISDEFAUT', False)+'"')
              else
                 EcritureEtablissement(FichierIE);
             EcritureModepaiement(FichierIE);

           // règlement
             AfficheListeCom('Réglements', TListBox(GetControl('LISTEEXPORT')));
             Listeregle := TList.Create;
             Chargementregle(Listeregle, '');
             Ecrituretableregle(Listeregle,FichierIE);
             Listeregle.Free ;
         end;
       // devise on n'envoie pas les devise cas balance
//       if Trans.balance <> TRUE then
       begin
         AfficheListeCom('Devise', TListBox(GetControl('LISTEEXPORT')));
         Ecrituredevise(FichierIE);
       end;

        if (Trans.Serie <> 'S1') then
        begin
       // régime TVA
         AfficheListeCom('TVA', TListBox(GetControl('LISTEEXPORT')));
         Ecrituretva(FichierIE);
        end;

       // section analytique
        if ((Trans.Pana = 'X') and (Trans.balance <> TRUE)) or ((GetControlText('BANA') = 'X') and (Trans.balance = TRUE)) then
        begin
         AfficheListeCom('Sections analytiques', TListBox(GetControl('LISTEEXPORT')));
         EcritureSectionana(FichierIE);
        end;
//       J_COMPTEURNORMAL
        if (Trans.Serie <> 'S1') and (Trans.balance <> TRUE) then
        begin
              // souche
             AfficheListeCom('Souches', TListBox(GetControl('LISTEEXPORT')));
             EcritureSouche(FichierIE,'');

           // Correspondance
             AfficheListeCom('Correspondants', TListBox(GetControl('LISTEEXPORT')));
             EcritureCorresp(FichierIE);
             EcritureCorrespimpot(FichierIE);
        end;
        // pour les comptes généraux

        if (Trans.Pgene = 'X')  or
        (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
        begin
          if(THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
           Where := ' Where G_DATECREATION >="'+ USDATETIME(GetParamSocSecur ('SO_CPRDDATERECEPTION', False)) +'"';

          AfficheListeCom('Comptes généraux', TListBox(GetControl('LISTEEXPORT')));
          if ListeCpteGen = nil then
          begin
               ListeCpteGen:=TList.Create;
               ChargementCpte(ListeCpteGen,Where);
          end;
          if (THValComboBox(GetControl('NATURETRANSFERT')).Value <> 'JRL') then
          begin
             EcritureCpptegen(ListeCpteGen,FichierIE);
             ListeCpteGen.Free ;
          end
          else
             EcritureCpptegen(ListeCpteGen,FichierIE, 'X');
          Where := '';
        end;

        // pour les comptes tiers
        if ((Trans.ptiers = 'X') and (Trans.balance <> TRUE)) or
        (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') or
        ((TCheckBox (GetControl ('AUXILIARE')).checked) and (Trans.balance = TRUE))then
        begin
          if(THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
          Where := ' Where T_DATECREATION >="'+ USDATETIME(GetParamSocSecur ('SO_CPRDDATERECEPTION', False)) +'"';
          AfficheListeCom('Comptes Tiers', TListBox(GetControl('LISTEEXPORT')));
          ListeCpteaux:=TList.Create;
          ListeRib:=TList.Create;
          ChargementCpteaux(ListeCpteaux,ListeRib,Where);
          if (THValComboBox(GetControl('NATURETRANSFERT')).Value <> 'JRL') then
          begin
             EcritureCppteaux(ListeCpteaux,FichierIE);
             ListeCpteaux.Free ;
             if (TOBTiers <> nil) and (stArg <> '') then TOBTiers.free;
          end
          else
             EcritureCppteaux(ListeCpteaux,FichierIE, 'X');
          // ajout me 23/05/2003
          EcritureRib(ListeRib,FichierIE);
          ListeRib.free;
          Where := '';
        end;

       // pour les journaux
          if Trans.balance = TRUE then
          begin
               if (Trans.Jr = '') or (Trans.Jr = '<<Tous>>')  then
               begin
                    if (ctxPCL in V_PGI.PGIContexte) then
                       Trans.Jr := 'CO'
                    else
                       Trans.Jr := 'CO;AA';
               end;
          end;

          ListeJalB:=TList.Create ;
          if Trans.Jr <> ''  then
          begin
               AfficheListeCom('Les journaux : ' + Trans.Jr, TListBox(GetControl('LISTEEXPORT')));
               Jourx := Trans.Jr;
               Where := ' Where ' + RendCommandeJournal('J',Jourx);
          end
          else
              AfficheListeCom('Les journaux', TListBox(GetControl('LISTEEXPORT')));
          AlimJalB(ListeJalB,Where) ;
          EcritureJournal(ListeJalB, FichierIE);
          ListeJalB.Free ;
          Where := '';

       // pour les écritures
 {
 WHERE E_DATECOMPTABLE>='01/01/2001'
 AND E_DATECOMPTABLE<='12/31/2001'
 AND E_EXERCICE='001'
 AND E_ETABLISSEMENT='166'
 AND E_NUMEROPIECE>=0 AND E_NUMEROPIECE<=999999999
 AND E_JOURNAL>='ACH' AND E_JOURNAL<='VEN' AND E_QUALIFPIECE='N' AND E_EXPORTE='X'
 }
        If (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') Then UpdateComLettrage(0) ;
        if (Trans.pecr = 'X')  and  (Trans.balance <> TRUE) then
        begin
          ListeMouvt:=TList.Create;
          RenWhereEcr(Where,'E');
          AfficheListeCom('Chargement des écritures', TListBox(GetControl('LISTEEXPORT')));
          if Trans.balance <> TRUE then
             ChargementMouvement(ListeMouvt,Where,Trans.TypeFormat,FichierIE, Trans.pana, ListeCpteGen, ListeCpteaux);
          AfficheListeCom('Export des écritures', TListBox(GetControl('LISTEEXPORT')));

          EcritureMouvement(ListeMouvt,FichierIE,Trans.TypeFormat);
          ListeMouvt.Free ;
        end;
        If (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') Then
        begin
// ajout me pour version 590
             if (THValComboBox(GetControl('HISTOSYNCHRO')).Value = '0') then
                UpdateComLettrage(4)
             else
             if (THValComboBox(GetControl('HISTOSYNCHRO')).Value = '1') then
                UpdateComLettrage(5)
             else
                UpdateComLettrage(3) ;
(*
                UpdateComLettrage(3) ;
*)
        end;
// ajout me pour version 590
        If (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') or (Trans.Serie = 'S1') Then
           SetParamsoc ('SO_CPDERNEXOCLO','');
        if (Trans.balance = TRUE) then
        begin
               AfficheListeCom('Balance', TListBox(GetControl('LISTEEXPORT')));
               ListeMouvt:=TList.Create;
               ChargementBalance(ListeMouvt, FichierIE);
               EcritureMouvement(ListeMouvt,FichierIE,Trans.TypeFormat);
               ListeMouvt.Free ;
        end;
        if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'JRL') then
        begin
            if  ListeCpteGen <> nil then
            begin
                 LibererGen (ListeCpteGen);
                 ListeCpteGen.Free ;
            end;
            if ListeCpteaux <> nil then
            begin
                 LibererAux (ListeCpteaux);
                 ListeCpteaux.Free;
                 if (TOBTiers <> nil) and (stArg <> '') then TOBTiers.free;
            end;
        end;
    finally
        CloseFile(FichierIE);
        _DeblocageMonoPoste(True) ;

    end ;
    AfficheListeCom('Export terminé ', TListBox(GetControl('LISTEEXPORT')));
// ajout me V590
    if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
       SetParamsoc ('SO_CPSYNCHROSX', TRUE);
    if (Trans.Serie = 'S1') and (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'DOS')
       and (GetParamSocSecur ('SO_CPDECLOTUREDETAIL', False)= TRUE) then
    SetParamsoc ('SO_CPDECLOTUREDETAIL', FALSE);

    if stArg <> '' then
    begin
         FTimer.free;
         Ecran.ModalResult := 1;
         Rep := ExtractFileDir (Trans.FichierSortie);
         FichierImp := Rep+'\'+'ListeCom'+FormatDateTime('yyyymmddhhnn',NowH)+'.txt';
    end
    else
        FichierImp := 'ListeCom'+FormatDateTime('yyyymmddhhnn',NowH)+'.txt';
    EcrireDansfichierListeCom (FichierImp, TListBox(GetControl('LISTEEXPORT')));
    if stArg = '' then
    begin
      if PGIAsk('Export terminé : voulez-vous imprimer le rapport ?','Export')=mrYes then
        ControlTextToPrinter(FichierImp,poPortrait);
    end
    else
    begin
         if Email <> '' then
         begin
             ListeEmail := TStringList.Create;
             ListeEmail.add ('Ci-joint le rapport de Comsx');
             SendMail('Rapport Comsx', Email, '', ListeEmail, FichierImp, true, 1);
             ListeEmail.free;
         end;
     end;
end;

Procedure TOF_ExportCom.AlimJalB(LJB : TList; Where : string) ;
Var
QJ        : TQuery ;
TJournal  : PListeJournal;
SWhere    : string;
BEGIN
if Where <> '' then SWhere := Where;
QJ:=OpenSQL('SELECT J_JOURNAL,J_LIBELLE,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,'+
 'J_CONTREPARTIE,J_MODESAISIE,J_COMPTEAUTOMAT,J_COMPTEINTERDIT,J_ABREGE,J_AXE FROM JOURNAL'+ sWhere,TRUE) ;
While Not QJ.Eof Do BEGIN
    System.New (TJournal);
    TJournal^.code := QJ.FindField ('J_JOURNAL').asstring;
    TJournal^.libelle := QJ.FindField ('J_LIBELLE').asstring;
    TJournal^.nature := QJ.FindField ('J_NATUREJAL').asstring;
    if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1') 
    and (TJournal^.nature = 'REG') then TJournal^.nature := 'OD';

    TJournal^.souchen := QJ.FindField ('J_COMPTEURNORMAL').asstring;
    TJournal^.souches := QJ.FindField ('J_COMPTEURSIMUL').asstring;
    TJournal^.compte := QJ.FindField ('J_CONTREPARTIE').asstring;
    TJournal^.axe := QJ.FindField ('J_AXE').asstring;
    TJournal^.modesaisie := QJ.FindField ('J_MODESAISIE').asstring;
    TJournal^.cptauto := QJ.FindField ('J_COMPTEAUTOMAT').asstring;
    TJournal^.cptint := QJ.FindField ('J_COMPTEINTERDIT').asstring;
    TJournal^.abrege := QJ.FindField ('J_ABREGE').asstring;
    LJB.Add(TJournal) ;

    QJ.NExt ;
END ;
if (Trans.balance = TRUE)  and QJ.EOF then
begin
    System.New (TJournal);
    TJournal^.code := 'CO';
    TJournal^.libelle := 'Reprise balance';
    TJournal^.nature := 'OD';
    TJournal^.souchen := 'CPT';
    TJournal^.souches := 'SIM';
    TJournal^.compte := '';
    TJournal^.modesaisie := 'LIB';
    TJournal^.cptauto := '';
    TJournal^.cptint := '';
    TJournal^.abrege := 'Reprise balance';
    LJB.Add(TJournal) ;
    if not (ctxPCL in V_PGI.PGIContexte) then
    begin
              System.New (TJournal);
              TJournal^.code := 'AA';
              TJournal^.libelle := 'Reprise balance Anouveau';
              TJournal^.nature := 'ANO';
              TJournal^.souchen := 'ANO';
              TJournal^.souches := 'SIM';
              TJournal^.compte := '';
              TJournal^.modesaisie := '-';
              TJournal^.cptauto := '';
              TJournal^.cptint := '';
              TJournal^.abrege := 'Reprise balance ANV';
              LJB.Add(TJournal) ;
    end;
end;

Ferme(QJ) ;

END ;

procedure TOF_ExportCom.ChargementTableLibre(ListeLibre : TList;Where : string);
var
Q1           :TQuery;
TLibre       : PLibre;
typecpte     : string;
begin
Q1:=OpenSQL('SELECT * FROM NATCPTE'+ Where,TRUE) ;
    While Not Q1.Eof Do BEGIN
      System.New (TLibre);
      TLibre^.identifiant := copy (Q1.FindField ('NT_TYPECPTE').asstring,3,1);
      typecpte := copy (Q1.FindField ('NT_TYPECPTE').asstring,1,1);
      TLibre^.code := Q1.FindField ('NT_NATURE').asstring;
      TLibre^.libelle := Q1.FindField ('NT_LIBELLE').asstring;

//T00 à T09 pour les tiers
//G00 à G09 pour les généraux
//S00 à S09 pour les sections
//E00 à E03 pour les ecritures

      if typecpte = 'T' then
         TLibre^.typetable := 'TIE' // pour tiers
      else
          if typecpte = 'G' then
                TLibre^.typetable := 'GEN' // pour généraux
      else
          if typecpte = 'S' then
                TLibre^.typetable := 'SEC' // pour section
      else
          if typecpte = 'E' then
                 TLibre^.typetable := 'ECR' // pour écriture
      else
          if typecpte = 'A' then
                 TLibre^.typetable := 'ANA' // pour analytique
      else
                 TLibre^.typetable := typecpte; // pour autres
      TLibre^.tx0 := Q1.FindField ('NT_TEXTE0').asstring;
      TLibre^.tx1 := Q1.FindField ('NT_TEXTE1').asstring;
      TLibre^.tx2 := Q1.FindField ('NT_TEXTE2').asstring;
      TLibre^.tx3 := Q1.FindField ('NT_TEXTE3').asstring;
      TLibre^.tx4 := Q1.FindField ('NT_TEXTE4').asstring;
      TLibre^.tx5 := Q1.FindField ('NT_TEXTE5').asstring;
      TLibre^.tx6 := Q1.FindField ('NT_TEXTE6').asstring;
      TLibre^.tx7 := Q1.FindField ('NT_TEXTE7').asstring;
      TLibre^.tx8 := Q1.FindField ('NT_TEXTE8').asstring;
      TLibre^.tx9 := Q1.FindField ('NT_TEXTE9').asstring;
      ListeLibre.Add(TLibre) ;
    Q1.NExt ;
   END ;
Ferme(Q1) ;
end;

procedure TOF_ExportCom.ChargementTableSection(ListeSection : TList;Where : string);
var
Q1,Q2        :TQuery;
TSection     :PSection;
begin
            Q2:=OpenSQL('SELECT PS_SOUSSECTION,PS_LIBELLE,PS_AXE,PS_CODE FROM SSSTRUCR',TRUE) ;
            while not Q2.EOF do
            begin
            System.New (TSection);
            TSection^.code := Q2.FindField ('PS_CODE').asstring;
            TSection^.libelle := Q2.FindField ('PS_LIBELLE').asstring;
            TSection^.axe := Q2.FindField ('PS_AXE').asstring;
            TSection^.codesp := Q2.FindField ('PS_SOUSSECTION').asstring;
//            TSection^.libellesp := Q2.FindField ('PS_LIBELLE').asstring;
                Q1:=OpenSQL('SELECT SS_AXE,SS_SOUSSECTION,SS_LIBELLE,SS_DEBUT,SS_LONGUEUR FROM STRUCRSE Where SS_SOUSSECTION="'+
                                    Q2.FindField ('PS_SOUSSECTION').asstring+'"',TRUE) ;
                if not Q1.EOF then
                begin
                TSection^.libellesp := Q1.FindField ('SS_LIBELLE').asstring;
                TSection^.debsp := Q1.FindField ('SS_DEBUT').asstring;
    //            TSection^.axe := Q1.FindField ('SS_AXE').asstring;
                TSection^.lgsp := IntToStr(Q1.FindField ('SS_LONGUEUR').asinteger);
                end;
                ferme (Q1);
                ListeSection.Add(TSection) ;
            Q2.NExt ;
            end;
            ferme (Q2);
end;

procedure TOF_ExportCom.ChargementCpte(ListeCpteGen : TList;Where : string);
var
Q1           :TQuery;
TCompteGene  : PListeCpteGen;
i, j : integer;
begin
 i:= gettickcount;
Q1:=OpenSQL('SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_LETTRABLE,G_POINTABLE,G_VENTILABLE1,'+
 'G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_TABLE0,'+
 'G_TABLE1,G_TABLE2,G_TABLE3,G_TABLE4,G_TABLE5,G_TABLE6,G_TABLE7,G_TABLE8,G_TABLE9,'
 +'G_ABREGE,G_SENS,G_CORRESP1,G_CORRESP2 FROM GENERAUX'+ Where,TRUE) ;
While Not Q1.Eof Do BEGIN
      System.New (TCompteGene);
      TCompteGene^.code        := Q1.FindField ('G_GENERAL').asstring;
      TCompteGene^.libelle     := Q1.FindField ('G_LIBELLE').asstring;
      TCompteGene^.nature      := Q1.FindField ('G_NATUREGENE').asstring;
      TCompteGene^.Lettrable   := Q1.FindField ('G_LETTRABLE').asstring;
      TCompteGene^.Pointage    := Q1.FindField ('G_POINTABLE').asstring;
      TCompteGene^.Ventilaxe1  := Q1.FindField ('G_VENTILABLE1').asstring;
      TCompteGene^.Ventilaxe2  := Q1.FindField ('G_VENTILABLE2').asstring;
      TCompteGene^.Ventilaxe3  := Q1.FindField ('G_VENTILABLE3').asstring;
      TCompteGene^.Ventilaxe4  := Q1.FindField ('G_VENTILABLE4').asstring;
      TCompteGene^.Ventilaxe5  := Q1.FindField ('G_VENTILABLE5').asstring;
      TCompteGene^.Table1      := Q1.FindField ('G_TABLE0').asstring;
      TCompteGene^.Table2      := Q1.FindField ('G_TABLE1').asstring;
      TCompteGene^.Table3      := Q1.FindField ('G_TABLE2').asstring;
      TCompteGene^.Table4      := Q1.FindField ('G_TABLE3').asstring;
      TCompteGene^.Table5      := Q1.FindField ('G_TABLE4').asstring;
      TCompteGene^.Table6      := Q1.FindField ('G_TABLE5').asstring;
      TCompteGene^.Table7      := Q1.FindField ('G_TABLE6').asstring;
      TCompteGene^.Table8      := Q1.FindField ('G_TABLE7').asstring;
      TCompteGene^.Table9      := Q1.FindField ('G_TABLE8').asstring;
      TCompteGene^.Table10     := Q1.FindField ('G_TABLE9').asstring;
      TCompteGene^.abrege      := Q1.FindField ('G_ABREGE').asstring;
      TCompteGene^.sens        := Q1.FindField ('G_SENS').asstring;
      // ajout me 13-05-2002
      TCompteGene^.Corresp1    := Q1.FindField ('G_CORRESP1').asstring;
      TCompteGene^.Corresp2    := Q1.FindField ('G_CORRESP2').asstring;
      ListeCpteGen.Add(TCompteGene) ;
    Q1.NExt ;
END ;
Ferme(Q1) ;
 j:= gettickcount;
// showmessage(inttostr(j-i));
end;

procedure TOF_ExportCom.ChargementCpteaux(ListeCpteaux,ListeRib : TList;Where : string);
var
Q1          : TQuery;
TCompteAux  : PListeAux;
Okecr       : Boolean;
ii          : integer;
begin
if (TOBTiers <> nil) and (stArg <> '') then
begin
    for ii :=0 to TOBTiers.detail.Count -1 do
    begin
            System.New (TCompteAux);
            TCompteAux^.code := TOBTiers.Detail[ii].getvalue('T_AUXILIAIRE');
            TCompteAux^.libelle := TOBTiers.Detail[ii].getvalue ('T_LIBELLE');
            TCompteAux.nature := TOBTiers.Detail[ii].getvalue('T_NATUREAUXI');
            TCompteAux^.lettrable  := TOBTiers.Detail[ii].getvalue('T_LETTRABLE');
            TCompteAux^.collectif := TOBTiers.Detail[ii].getvalue ('T_COLLECTIF');
            TCompteAux^.ean := TOBTiers.Detail[ii].getvalue ('T_EAN');
            TCompteAux^.tb1 := TOBTiers.Detail[ii].getvalue ('T_TABLE0');
            TCompteAux^.tb2 := TOBTiers.Detail[ii].getvalue ('T_TABLE1');
            TCompteAux^.tb3 := TOBTiers.Detail[ii].getvalue ('T_TABLE2');
            TCompteAux^.tb4 := TOBTiers.Detail[ii].getvalue ('T_TABLE3');
            TCompteAux^.tb5 := TOBTiers.Detail[ii].getvalue ('T_TABLE4');
            TCompteAux^.tb6 := TOBTiers.Detail[ii].getvalue ('T_TABLE5');
            TCompteAux^.tb7 := TOBTiers.Detail[ii].getvalue ('T_TABLE6');
            TCompteAux^.tb8 := TOBTiers.Detail[ii].getvalue ('T_TABLE7');
            TCompteAux^.tb9 := TOBTiers.Detail[ii].getvalue ('T_TABLE8');
            TCompteAux^.tb10 := TOBTiers.Detail[ii].getvalue ('T_TABLE9');
            TCompteAux^.adresse1 := TOBTiers.Detail[ii].getvalue('T_ADRESSE1');
            TCompteAux^.adresse2 := TOBTiers.Detail[ii].getvalue ('T_ADRESSE2');
            TCompteAux^.adresse3 := TOBTiers.Detail[ii].getvalue ('T_ADRESSE3');
            TCompteAux^.codepostal := TOBTiers.Detail[ii].getvalue ('T_CODEPOSTAL');
            TCompteAux^.ville := TOBTiers.Detail[ii].getvalue ('T_VILLE');
            TCompteAux^.pays := TOBTiers.Detail[ii].getvalue ('T_PAYS');
            TCompteAux^.abrege := TOBTiers.Detail[ii].getvalue ('T_ABREGE');
            TCompteAux^.langue := TOBTiers.Detail[ii].getvalue ('T_LANGUE');
            TCompteAux^.multidevise := TOBTiers.Detail[ii].getvalue ('T_MULTIDEVISE');
            TCompteAux^.devisetiers := TOBTiers.Detail[ii].getvalue ('T_DEVISE');
            TCompteAux^.tel := TOBTiers.Detail[ii].getvalue ('T_TELEPHONE');
            TCompteAux^.fax := TOBTiers.Detail[ii].getvalue('T_FAX');
            TCompteAux^.regimetva := TOBTiers.Detail[ii].getvalue ('T_REGIMETVA');
            TCompteAux^.modereg := TOBTiers.Detail[ii].getvalue('T_MODEREGLE');
            TCompteAux^.commentaire := TOBTiers.Detail[ii].getvalue ('T_COMMENTAIRE');
            TCompteAux^.nif := TOBTiers.Detail[ii].getvalue('T_NIF');
            TCompteAux^.siret := TOBTiers.Detail[ii].getvalue ('T_SIRET');
            TCompteAux^.ape := TOBTiers.Detail[ii].getvalue('T_APE');
            RechContact(TCompteAux^.code, TCompteAux);
            TCompteAux^.formjur := TOBTiers.Detail[ii].getvalue ('T_FORMEJURIDIQUE');
            TCompteAux^.tvaenc :=  TOBTiers.Detail[ii].getvalue ('T_TVAENCAISSEMENT');
                  // pour les ribs on écrit autant de fois que nombre de rib
            Okecr := RechRib (TCompteAux^.code, TCompteAux, ListeCpteaux,ListeRib);
            if not Okecr then
            ListeCpteaux.Add(TCompteAux) ;
    end;
    exit;
end;
//TOBTiers := TOB.Create('', nil, -1);

Q1:=OpenSQL('SELECT * FROM TIERS'+ Where,TRUE) ;
//TOBTiers.LoadDetailDB('TIERS', '', '', Q1, TRUE, FALSE);
//TOBTiers.SaveToFile('journal.txt',True,True,True);
//TOBTiers.free;
    While Not Q1.Eof Do BEGIN
      System.New (TCompteAux);
      TCompteAux^.code := Q1.FindField ('T_AUXILIAIRE').asstring;
      TCompteAux^.libelle := Q1.FindField ('T_LIBELLE').asstring;
      TCompteAux.nature := Q1.FindField ('T_NATUREAUXI').asstring;
      TCompteAux^.lettrable  := Q1.FindField ('T_LETTRABLE').asstring;
      TCompteAux^.collectif := Q1.FindField ('T_COLLECTIF').asstring;
      TCompteAux^.ean := Q1.FindField ('T_EAN').asstring;
      TCompteAux^.tb1 := Q1.FindField ('T_TABLE0').asstring;
      TCompteAux^.tb2 := Q1.FindField ('T_TABLE1').asstring;
      TCompteAux^.tb3 := Q1.FindField ('T_TABLE2').asstring;
      TCompteAux^.tb4 := Q1.FindField ('T_TABLE3').asstring;
      TCompteAux^.tb5 := Q1.FindField ('T_TABLE4').asstring;
      TCompteAux^.tb6 := Q1.FindField ('T_TABLE5').asstring;
      TCompteAux^.tb7 := Q1.FindField ('T_TABLE6').asstring;
      TCompteAux^.tb8 := Q1.FindField ('T_TABLE7').asstring;
      TCompteAux^.tb9 := Q1.FindField ('T_TABLE8').asstring;
      TCompteAux^.tb10 := Q1.FindField ('T_TABLE9').asstring;
      TCompteAux^.adresse1 := Q1.FindField ('T_ADRESSE1').asstring;
      TCompteAux^.adresse2 := Q1.FindField ('T_ADRESSE2').asstring;
      TCompteAux^.adresse3 := Q1.FindField ('T_ADRESSE3').asstring;
      TCompteAux^.codepostal := Q1.FindField ('T_CODEPOSTAL').asstring;
      TCompteAux^.ville := Q1.FindField ('T_VILLE').asstring;
      TCompteAux^.pays := Q1.FindField ('T_PAYS').asstring;
      TCompteAux^.abrege := Q1.FindField ('T_ABREGE').asstring;
      TCompteAux^.langue := Q1.FindField ('T_LANGUE').asstring;
      TCompteAux^.multidevise := Q1.FindField ('T_MULTIDEVISE').asstring;
      TCompteAux^.devisetiers := Q1.FindField ('T_DEVISE').asstring;
      TCompteAux^.tel := Q1.FindField ('T_TELEPHONE').asstring;
      TCompteAux^.fax := Q1.FindField ('T_FAX').asstring;
      TCompteAux^.regimetva := Q1.FindField ('T_REGIMETVA').asstring;
      TCompteAux^.modereg := Q1.FindField ('T_MODEREGLE').asstring;
      TCompteAux^.commentaire := Q1.FindField ('T_COMMENTAIRE').asstring;
      TCompteAux^.nif := Q1.FindField ('T_NIF').asstring;
      TCompteAux^.siret := Q1.FindField ('T_SIRET').asstring;
      TCompteAux^.ape := Q1.FindField ('T_APE').asstring;
      RechContact(TCompteAux^.code, TCompteAux);
      TCompteAux^.formjur := Q1.FindField ('T_FORMEJURIDIQUE').asstring;
      TCompteAux^.tvaenc :=  Q1.FindField ('T_TVAENCAISSEMENT').asstring;
            // pour les ribs on écrit autant de fois que nombre de rib
      Okecr := RechRib (TCompteAux^.code, TCompteAux, ListeCpteaux,ListeRib);
      if not Okecr then
      ListeCpteaux.Add(TCompteAux) ;
    Q1.NExt ;
   END ;
Ferme(Q1) ;
end;

FUNCTION TOF_ExportCom.RechRib(cpteaux  : string; var TCompteAux  : PListeAux; var ListeCpteaux,ListeRib : TList) : Boolean;
var Qr        : Tquery;
TRib          : PListeRib;
begin
Result := FALSE;
    Qr := OpenSql ('SELECT * FROM RIB'+
    ' Where R_AUXILIAIRE="'+cpteaux+'"', TRUE);
    while not Qr.EOF do
    begin
      System.New (TRib);
     if  Qr.FindField ('R_NUMERORIB').asinteger > 1 then
         System.New (TCompteAux);
     TCompteAux^.domicile := Qr.FindField ('R_DOMICILIATION').asstring;
     TCompteAux^.etablissement := Qr.FindField ('R_ETABBQ').asstring;
     TCompteAux^.guichet := Qr.FindField ('R_GUICHET').asstring;
     TCompteAux^.compte := Qr.FindField ('R_NUMEROCOMPTE').asstring;
     TCompteAux^.cle := Qr.FindField ('R_CLERIB').asstring;
     TCompteAux^.rib := Qr.FindField ('R_PRINCIPAL').asstring;
     ListeCpteaux.Add (TCompteAux);
// ajout me  23/05/2003
     TRib^.ident := 'RIB';
     TRib^.code  := cpteaux;
     Trib^.numero := IntToStr(Qr.FindField ('R_NUMERORIB').asinteger);
     Trib^.prinpipal := Qr.FindField ('R_PRINCIPAL').asstring;
     Trib^.codebanque :=  Qr.FindField ('R_ETABBQ').asstring;
     Trib^.codeguichet := Qr.FindField ('R_GUICHET').asstring;
     Trib^.nocompte := Qr.FindField ('R_NUMEROCOMPTE').asstring;
     Trib^.cle := Qr.FindField ('R_CLERIB').asstring;
     Trib^.domicile := Qr.FindField ('R_DOMICILIATION').asstring;
     Trib^.ville := Qr.FindField ('R_VILLE').asstring;
     Trib^.pays := Qr.FindField ('R_PAYS').asstring;
     Trib^.devise := Qr.FindField ('R_DEVISE').asstring;
     Trib^.bic := Qr.FindField ('R_CODEBIC').asstring;
     Trib^.soc := Qr.FindField ('R_SOCIETE').asstring;
     Trib^.ribsal := Qr.FindField ('R_SALAIRE').asstring;
     Trib^.ribcompte := Qr.FindField ('R_ACOMPTE').asstring;
     Trib^.ribprof := Qr.FindField ('R_FRAISPROF').asstring;
     ListeRib.Add (Trib);
     Result := TRUE;
     Qr.next;
    end;
ferme (Qr);
end;

procedure TOF_ExportCom.RechContact(cpteaux  : string; var TCompteAux  : PListeAux);
var Qc    : Tquery;
begin
    Qc := OpenSql ('SELECT  C_NOM,C_SERVICE,C_FONCTION,C_TELEPHONE,C_FAX,C_TELEX,C_RVA,C_CIVILITE,C_PRINCIPAL FROM CONTACT'+
    ' Where C_AUXILIAIRE="'+cpteaux+'" AND C_TYPECONTACT="T"', TRUE);
    if not Qc.EOF then
    begin
     TCompteAux^.ctnom := Qc.FindField ('C_NOM').asstring;
     TCompteAux^.ctservice := Qc.FindField ('C_SERVICE').asstring;
     TCompteAux^.ctfonction := Qc.FindField ('C_FONCTION').asstring;
     TCompteAux^.cttel := Qc.FindField ('C_TELEPHONE').asstring;
     TCompteAux^.ctfax := Qc.FindField ('C_FAX').asstring;
     TCompteAux^.cttelex := Qc.FindField ('C_TELEX').asstring;
     TCompteAux^.ctrva := Qc.FindField ('C_RVA').asstring;
     TCompteAux^.ctcivilite := Qc.FindField ('C_CIVILITE').asstring;
     TCompteAux^.ctprincipal := Qc.FindField ('C_PRINCIPAL').asstring;
    end;
    ferme (Qc);
end;

procedure TOF_ExportCom.ChargementMouvement(var ListeMouvt : TList;Where,FormatEnvoie : string;Var F : TextFile; pana : string; ListeCpteGen : TList; ListeCpteaux : TList);
var
Q1                            :TQuery;
TMvt                          : PListeMouvt;
i, j                          : integer;
OkEuro                        : Boolean;
DecDev                        : integer ;
CodeMontant                   : string;
EnDevise,EnMontantOppose      : Boolean ;
Quotite                       : Double;
TDevise                       : TFDevise ;
TDev                          : TList;
SaisieOpposee                 : Boolean ;
Orderby                       : string;
Wheres,Verrou,VerrouS5,Whereana : string;
AncJournal                      : string;
Ancnumpiece                     : integer;
Ancexercice, Ancienperiode      : string;
OkLett                          : string;
begin
 AncJournal := ''; Ancnumpiece := 0;
 Orderby := ' order by E_JOURNAL, E_EXERCICE, E_PERIODE, E_NUMEROPIECE, E_QUALIFPIECE,E_NUMLIGNE,E_NUMECHE';
 if Where <> '' then
 begin
      if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
      begin
// ajout me pour version 590
             if (THValComboBox(GetControl('HISTOSYNCHRO')).Value = '0') then
                Wheres := 'Where '+ Where + ' AND E_QUALIFPIECE="N" AND E_IO="0"'
             else
             if (THValComboBox(GetControl('HISTOSYNCHRO')).Value = '1') then
                Wheres := 'Where '+ Where + ' AND E_QUALIFPIECE="N" AND E_IO="1"'
             else
                Wheres := 'Where '+ Where + ' AND E_QUALIFPIECE="N" AND E_IO="X"'
      end
      else
         // ajout me 15-07-2003 Wheres := 'Where '+ Where + ' AND E_QUALIFPIECE="N"';
         Wheres := 'Where '+ Where ;
 end;
 i:= gettickcount;
 DecDev:=V_PGI.OkDecV ;
 Q1:=OpenSQL('SELECT D_DEVISE,D_DECIMALE,D_QUOTITE FROM DEVISE ORDER BY D_DEVISE',True) ;
 TDev:=TList.Create ;
 While not Q1.Eof do
  BEGIN
  TDevise:=TFDevise.Create ;
  TDevise.Code:=Q1.Fields[0].AsString ;
  TDevise.Decimale:=Q1.Fields[1].AsInteger ;
  TDevise.Quotite:=Q1.Fields[2].AsFloat ;
  TDev.Add(TDevise) ;
  Q1.Next ;
  END ;
  Ferme(Q1) ;

  // prendre dans la table devise le decimal pour les arrondi

Q1:=OpenSQL('SELECT * FROM ECRITURE '+ Wheres + OrderBy,TRUE);
if Q1.EOF then
begin
  if stArg <> '' then  // par ligne de commande
           AfficheListeCom('Ecriture du journal : ',TListBox(GetControl('LISTEEXPORT')))
  else
  begin
     PGIInfo ('Les critères de sélection ne renvoient aucun enregistrement.','Export');
     Ferme(Q1) ; exit;
  end;
end;
While Not Q1.Eof Do BEGIN
// écriture des dans le fichier dès qu'il y a une rupture E_JOURNAL, E_EXERCICE, E_PERIODE, E_NUMEROPIECE

      if AncJournal <> Q1.FindField ('E_JOURNAL').asstring then
      begin
           EcritureMouvement(ListeMouvt,F,Trans.TypeFormat);
           ListeMouvt.Free ;
           ListeMouvt:=TList.Create;
           AfficheListeCom('Ecriture du journal : '+ Q1.FindField ('E_JOURNAL').asstring,TListBox(GetControl('LISTEEXPORT')));
      end;
      AncJournal := Q1.FindField ('E_JOURNAL').asstring;

      if Ancexercice <> Q1.FindField ('E_EXERCICE').asstring then
      begin
           EcritureMouvement(ListeMouvt,F,Trans.TypeFormat);
           ListeMouvt.Free ;
           ListeMouvt:=TList.Create;
      end;
      Ancexercice := Q1.FindField ('E_EXERCICE').asstring;

      if (AncienPeriode <> Q1.FindField ('E_PERIODE').asstring) then
      begin
           EcritureMouvement(ListeMouvt,F,Trans.TypeFormat);
           ListeMouvt.Free ;
           ListeMouvt:=TList.Create;
      end;
      AncienPeriode := Q1.FindField ('E_PERIODE').asstring;

      if (Ancnumpiece <> Q1.FindField ('E_NUMEROPIECE').asinteger) then
      begin
           EcritureMouvement(ListeMouvt,F,Trans.TypeFormat);
           ListeMouvt.Free ;
           ListeMouvt:=TList.Create;
      end;
      Ancnumpiece := Q1.FindField ('E_NUMEROPIECE').asinteger;

      System.New (TMvt);
      TMvt^.Journal := Q1.FindField ('E_JOURNAL').asstring;
      TMvt^.Datecomptable := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATECOMPTABLE').AsDateTime);
      TMvt^.Typepiece := Q1.FindField ('E_NATUREPIECE').asstring;
      TMvt^.General := Q1.FindField ('E_GENERAL').asstring;
      TMvt^.AuxSect := Q1.FindField ('E_AUXILIAIRE').asstring;
      if TMvt^.AuxSect <>'' then TMvt^.Typecpte:= 'X' else TMvt^.Typecpte:=' ' ;
      TMvt^.Refinterne := Q1.FindField ('E_REFINTERNE').asstring;
      TMvt^.Libelle := Q1.FindField ('E_LIBELLE').asstring;
      TMvt^.Modepaie := Q1.FindField ('E_MODEPAIE').asstring;
      TMvt^.Echeance := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATEECHEANCE').AsDateTime);

      if (((arrondi (Q1.FindField ('E_DEBIT').AsFloat, V_PGI.OKDECV)) > 0) or
         ((arrondi(Q1.FindField ('E_DEBIT').AsFloat,V_PGI.OKDECV)) < 0) and
         ((arrondi (Q1.FindField ('E_CREDIT').AsFloat,V_PGI.OKDECV)) =0))
      then
           TMvt^.Sens := 'D';

      if (((Q1.FindField ('E_CREDIT').AsFloat > 0) or
      (Q1.FindField ('E_CREDIT').AsFloat < 0)) or (Q1.FindField ('E_CREDIT').AsFloat>0)) and (Q1.FindField ('E_DEBIT').AsFloat=0) then
           TMvt^.Sens := 'C';

      TMvt^.Typeecriture := Q1.FindField ('E_QUALIFPIECE').asstring;
      TMvt^.NumPiece := IntToStr(Q1.FindField ('E_NUMEROPIECE').asinteger);
      TMvt^.Devise := Q1.FindField ('E_DEVISE').asstring;
      TMvt^.TauxDev := AlignDroite(StrfMontant(Q1.FindField ('E_COTATION').asfloat,20,9,'',False),10);

      OkEuro:=Q1.FindField('E_DATECOMPTABLE').AsDateTime>=V_PGI.DateDebutEuro ;
      If OkEuro Then
         CodeMontant := 'FDE'
      else
          CodeMontant := 'FD-';
      EnDevise:=RecupDevise(TMvt^.Devise,DecDev,Quotite,TDev) ;
      If EnDevise Then
      BEGIN
           CodeMontant:='DFE' ;
           TMvt^.Montant1:=QuelMontant(Q1,'E',1,DecDev,TMvt^.Sens) ;
           TMvt^.Montant2:=QuelMontant(Q1,'E',0,V_PGI.OkDecV,TMvt^.Sens) ;
           TMvt^.Montant3:=QuelMontant(Q1,'E',2,V_PGI.OkDecE,TMvt^.Sens) ;
           If VH^.TenueEuro Then CodeMontant:='DEF' ;
      END Else
      BEGIN
//           If OkEuro Then
           BEGIN
                EnMontantOppose:=Q1.FindField('E_SAISIEEURO').AsString='X' ;
                If EnMontantOppose Then
                BEGIN
                     If VH^.TenueEuro Then
                          CodeMontant:='FE-'
                     Else
                          CodeMontant:='EF-' ;
                     TMvt^.Montant2:=QuelMontant(Q1,'E',0,V_PGI.OkDecV,TMvt^.Sens) ;
                     TMvt^.Montant1:=QuelMontant(Q1,'E',2,V_PGI.OkDecE,TMvt^.Sens) ;
                     TMvt^.Montant3:=Format_String(' ',20) ;
                END Else
                BEGIN
                      If VH^.TenueEuro Then
                           CodeMontant:='EF-'
                      Else
                           CodeMontant:='FE-' ;
                      TMvt^.Montant1:=QuelMontant(Q1,'E',0,V_PGI.OkDecV,TMvt^.Sens) ;
                      TMvt^.Montant2:=QuelMontant(Q1,'E',2,V_PGI.OkDecE,TMvt^.Sens) ;
                      TMvt^.Montant3:=Format_String(' ',20) ;
                END ;
           END;
(*
           Else
           BEGIN
               CodeMontant:='FE-' ; // à revoir pour 1998
               TMvt^.Montant1:=QuelMontant(Q1,'E',0,V_PGI.OkDecV,TMvt^.Sens) ;
               TMvt^.Montant2:=QuelMontant(Q1,'E',2,V_PGI.OkDecE,TMvt^.Sens) ;
               TMvt^.Montant3:=Format_String(' ',20) ;
           END ;
*)
     END ;

   TMvt^.CodeMontant := Codemontant;
   if (EtablissUnique) and (stArg <> '') then
      TMvt^.Etablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', False)
   else
      TMvt^.Etablissement := Q1.FindField ('E_ETABLISSEMENT').asstring;
   TMvt^.Axe := '';

   TMvt^.Numeche := IntToStr(Q1.FindField ('E_NUMECHE').asinteger);
// traitement des devises
  If OkEuro Then
     BEGIN
     SaisieOpposee:=Q1.FindField('E_SAISIEEURO').AsString='X' ;
     If TMvt^.Devise=V_PGI.DevisePivot Then
        BEGIN
        If ((VH^.TenueEuro) And (Not SaisieOpposee)) Or
           ((Not VH^.TenueEuro) And (SaisieOpposee)) Then TMvt^.Devise:='EUR' ;
        END ;
     END;
     if FormatEnvoie = 'ETE' then // format étendu
     begin
         TMvt^.RefExterne := Q1.findField('E_REFEXTERNE').Asstring;
         TMvt^.Daterefexterne := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATEREFEXTERNE').AsDateTime);
         TMvt^.Datecreation := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATECREATION').AsDateTime);
         TMvt^.Societe := Q1.findField('E_SOCIETE').Asstring;
         TMvt^.Affaire := Q1.findField('E_AFFAIRE').Asstring;
         TMvt^.Datetauxdev := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATETAUXDEV').AsDateTime);
         TMvt^.Ecranouveau := Q1.findField('E_ECRANOUVEAU').Asstring;
         TMvt^.Qte1 := AlignDroite(StrfMontant(Q1.FindField('E_QTE1').AsFloat,20,4,'',False),20);
         TMvt^.Qte2 := AlignDroite(StrfMontant(Q1.FindField('E_QTE2').AsFloat,20,4,'',False),20);
         TMvt^.QualifQte1 := Q1.findField('E_QUALIFQTE1').Asstring;
         TMvt^.QualifQte2 := Q1.findField('E_QUALIFQTE2').Asstring;
         TMvt^.Reflibre := Q1.findField('E_REFLIBRE').Asstring;
         TMvt^.Tvaencaiss := Q1.findField('E_TVAENCAISSEMENT').Asstring;
         TMvt^.Regimetva := Q1.findField('E_REGIMETVA').Asstring;
         TMvt^.Tva := Q1.findField('E_TVA').Asstring;
         TMvt^.TPF := Q1.findField('E_TPF').Asstring;
         TMvt^.Contrepartigen := Q1.findField('E_CONTREPARTIEGEN').Asstring;
         TMvt^.contrepartiaux := Q1.findField('E_CONTREPARTIEAUX').Asstring;
         TMvt^.RefPointage := Q1.findField('E_REFPOINTAGE').Asstring;
         TMvt^.datepointage := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATEPOINTAGE').AsDateTime);
         TMvt^.daterelance := FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATERELANCE').AsDateTime);
         TMvt^.datevaleur :=  FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATEVALEUR').AsDateTime);
         TMvt^.Rib := Q1.findField('E_RIB').Asstring;
         TMvt^.Refreleve := Q1.findField('E_REFRELEVE').Asstring;
         if (THValComboBox(GetControl('NATURETRANSFERT')).Value <> 'JRL') then
            TMvt^.Numeroimmo := Q1.findField('E_IMMO').Asstring
         else
            TMvt^.Numeroimmo := '';
         TMvt^.LT0 := Q1.findField('E_LIBRETEXTE0').Asstring;
         TMvt^.LT1 := Q1.findField('E_LIBRETEXTE1').Asstring;
         TMvt^.LT2 := Q1.findField('E_LIBRETEXTE2').Asstring;
         TMvt^.LT3 := Q1.findField('E_LIBRETEXTE3').Asstring;
         TMvt^.LT4 := Q1.findField('E_LIBRETEXTE4').Asstring;
         TMvt^.LT5 := Q1.findField('E_LIBRETEXTE5').Asstring;
         TMvt^.LT6 := Q1.findField('E_LIBRETEXTE6').Asstring;
         TMvt^.LT7 := Q1.findField('E_LIBRETEXTE7').Asstring;
         TMvt^.LT8 := Q1.findField('E_LIBRETEXTE8').Asstring;
         if (THRadioGroup(GetControl('TRANSFERTVERS')).Value = 'S1')
         or (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
         begin
            if Q1.findField('E_ETATREVISION').Asstring = 'X' then
               Verrou := 'X'
            else
               Verrou := '-';
            if Q1.findField('E_PAQUETREVISION').Asinteger = 1 then
               VerrouS5 := 'X'
            else
               VerrouS5 := '-';
            if Q1.findField('E_REFREVISION').Asinteger = 0 then
               TMvt^.LT9 := Format ('      %1.1s%-3.3s%-6.6s%-.06d%1.1s%-4.4s',
               [Verrou, TMvt^.Journal, Q1.findField('E_PERIODE').Asstring, Q1.FindField ('E_NUMEROPIECE').asinteger, VerrouS5,
               // ajout me 24-05-2002 pour lettrage s1
               Copy (Q1.findField('E_LIBRETEXTE9').Asstring,24,4)])
            else
                TMvt^.LT9 := Format ('%.06d%1.1s%-3.3s%-6.6s%-.06d%1.1s%-4.4s',
                [Q1.findField('E_REFREVISION').Asinteger, Verrou, TMvt^.Journal,
                Q1.findField('E_PERIODE').Asstring, Q1.FindField ('E_NUMEROPIECE').asinteger, VerrouS5,
               // ajout me 24-05-2002 pour lettrage s1
                Copy (Q1.findField('E_LIBRETEXTE9').Asstring,24,4)]);
         end
         else
             TMvt^.LT9 := Q1.findField('E_LIBRETEXTE9').Asstring;
         TMvt^.TA0 := Q1.findField('E_TABLE0').Asstring;
         TMvt^.TA1 := Q1.findField('E_TABLE1').Asstring;
         TMvt^.TA2 := Q1.findField('E_TABLE2').Asstring;
         TMvt^.TA3 := Q1.findField('E_TABLE3').Asstring;
         TMvt^.LM0 := Q1.findField('E_LIBREMONTANT0').Asstring;
         TMvt^.LM1 := Q1.findField('E_LIBREMONTANT1').Asstring;
         TMvt^.LM2 := Q1.findField('E_LIBREMONTANT2').Asstring;
         TMvt^.LM3 := Q1.findField('E_LIBREMONTANT3').Asstring;
         TMvt^.LD :=  FormatDateTime('ddmmyyyy',Q1.FindField ('E_LIBREDATE').AsDateTime);
         TMvt^.LB0 := Q1.findField('E_LIBREBOOL0').Asstring;
         TMvt^.LB1 := Q1.findField('E_LIBREBOOL1').Asstring;
         TMvt^.Conso := Q1.findField('E_CONSO').Asstring;
         TMvt^.Datepaquetmax :=  FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATEPAQUETMAX').AsDateTime);
         TMvt^.Datepaquetmin :=  FormatDateTime('ddmmyyyy',Q1.FindField ('E_DATEPAQUETMIN').AsDateTime);
         // on enleve le lettrage dans le cas du journal
         if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'JRL') then
         begin
             TMvt^.Couverture := AlignDroite(StrfMontant(0,20,V_PGI.OkDecV,'',False),20);
             TMvt^.Couverturedev := AlignDroite(StrfMontant(0,20,V_PGI.OkDecV,'',False),20);
             TMvt^.Couvertureeuro :=AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20);
             TMvt^.lettrage := '';
             TMvt^.lettragedev := '-';
             TMvt^.lettrageeuro := '-';
             if TMvt^.Typecpte = 'X' then
                OkLett := OkLettrableCpptetiers(ListeCpteaux, TMvt^.AuxSect)
             else
                 OkLett := OkLettrableCpptegen(ListeCpteGen, TMvt^.General);
             if OkLett <> 'X' then
                TMvt^.etatlettrage := 'RI'
             else
                TMvt^.etatlettrage := 'AL';
             TMvt^.RefPointage := '';
             TMvt^.datepointage := FormatDateTime('ddmmyyyy',iDate1900);
         end
         else
         begin
             TMvt^.Couverture := AlignDroite(StrfMontant(Q1.FindField('E_COUVERTURE').AsFloat,20,V_PGI.OkDecV,'',False),20);
             TMvt^.Couverturedev := AlignDroite(StrfMontant(Q1.FindField('E_COUVERTUREDEV').AsFloat,20,V_PGI.OkDecV,'',False),20);
             TMvt^.Couvertureeuro :=AlignDroite(StrfMontant(Q1.FindField('E_COUVERTUREEURO').AsFloat,20,V_PGI.OkDecE,'',False),20);
             TMvt^.lettrage := Q1.findField('E_LETTRAGE').Asstring;
             TMvt^.lettragedev := Q1.findField('E_LETTRAGEDEV').Asstring;
             TMvt^.lettrageeuro := Q1.findField('E_LETTRAGEEURO').Asstring;
             TMvt^.etatlettrage := Q1.findField('E_ETATLETTRAGE').Asstring;
         end;
     end;
   ListeMouvt.Add(TMvt) ;

   if (Q1.FindField('E_ANA').AsString='X') And (Pana = 'X' ) then // traitement analytique
   begin
        Whereana := 'Y_JOURNAL="'+TMvt^.Journal+'" AND Y_DATECOMPTABLE="'+
        USDateTime (Q1.findField('E_DATECOMPTABLE').AsDatetime) +
        '" AND Y_NUMEROPIECE=' + TMvt^.NumPiece +
        ' AND Y_NUMLIGNE='+ IntToStr(Q1.findField('E_NUMLIGNE').Asinteger);
        ChargementMouvementAnalytique(ListeMouvt,Whereana,FormatEnvoie, Tdev, F);
   end;

   Q1.NExt ;
END ;
Ferme(Q1) ;
// pour les OD analytiques a voir od analytique pour synchro
if (Trans.Serie <> 'S1') then
begin
     Whereana := '';
     RenWhereEcr(Whereana,'Y');
     if Whereana <> '' then Whereana := Whereana + ' AND Y_TYPEANALYTIQUE="X"'
     else Whereana := 'Where Y_TYPEANALYTIQUE="X"';
     ChargementMouvementAnalytique(ListeMouvt,Whereana,FormatEnvoie, Tdev, F);
end;
if Where <> '' then Wheres := 'Where '+ Where;
//ExecuteSql('UPDATE ECRITURE SET E_EXPORTE="X" '+Wheres) ;

 j:= gettickcount;
// showmessage(inttostr(j-i));
if TDev <> nil then begin VideListe(TDev); TDev.Free ; end;
end;

procedure TOF_ExportCom.ChargementMouvementAnalytique(var ListeMouvt : TList;Where,FormatEnvoie : string; Tdev : TList;Var F : TextFile);
var
Q1                            :TQuery;
TMvt                          : PListeMouvt;
i, j                          : integer;
OkEuro                        : Boolean;
DecDev                        : integer ;
CodeMontant                   : string;
EnDevise,EnMontantOppose      : Boolean ;
Quotite                       : Double;
SaisieOpposee                 : Boolean ;
Orderby                       : string;
Wheres                        : string;
begin

 Orderby := 'ORDER BY Y_JOURNAL,Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_QUALIFPIECE,Y_AXE, Y_NUMVENTIL';
 if Where <> '' then
 begin
         Wheres := 'Where '+ Where + ' AND Y_QUALIFPIECE="N"';
 end;
 i:= gettickcount;
 DecDev:=V_PGI.OkDecV ;

  // prendre dans la table devise le decimal pour les arrondi

if FormatEnvoie = 'STD' then
Q1:=OpenSQL('SELECT Y_JOURNAL,Y_DATECOMPTABLE,Y_NATUREPIECE, Y_GENERAL,Y_QUALIFPIECE,'+
'Y_SECTION,Y_REFINTERNE,Y_LIBELLE,Y_DEBIT,Y_CREDIT,'+
'Y_NUMEROPIECE,Y_DEVISE,Y_TAUXDEV,Y_SAISIEEURO,Y_ETABLISSEMENT,Y_DEBIT,Y_CREDIT,Y_DEBITEURO,Y_CREDITEURO, '+
'Y_DEBITDEV,Y_CREDITDEV, Y_AXE, Y_TYPEANALYTIQUE FROM ANALYTIQ '+ Wheres + Orderby,TRUE)
else
Q1:=OpenSQL('SELECT * FROM ANALYTIQ '+ Wheres + OrderBy,TRUE);


While Not Q1.Eof Do BEGIN
      System.New (TMvt);

      TMvt^.Journal := Q1.FindField ('Y_JOURNAL').asstring;
      TMvt^.Datecomptable := FormatDateTime('ddmmyyyy',Q1.FindField ('Y_DATECOMPTABLE').AsDateTime);
      TMvt^.Typepiece := Q1.FindField ('Y_NATUREPIECE').asstring;
      TMvt^.General := Q1.FindField ('Y_GENERAL').asstring;
      TMvt^.AuxSect := Q1.FindField ('Y_SECTION').asstring;
      if Q1.FindField ('Y_TYPEANALYTIQUE').asstring = 'X' then
         TMvt^.Typecpte:='O'
      else
         TMvt^.Typecpte:='A' ;
      TMvt^.Refinterne := Q1.FindField ('Y_REFINTERNE').asstring;
      TMvt^.Libelle := Q1.FindField ('Y_LIBELLE').asstring;
      TMvt^.Echeance := '        ';

      if (((arrondi (Q1.FindField ('Y_DEBIT').AsFloat, V_PGI.OKDECV)) > 0) or
         ((arrondi(Q1.FindField ('Y_DEBIT').AsFloat,V_PGI.OKDECV)) < 0) and
         ((arrondi (Q1.FindField ('Y_CREDIT').AsFloat,V_PGI.OKDECV)) =0))
      then
           TMvt^.Sens := 'D';

      if (((Q1.FindField ('Y_CREDIT').AsFloat > 0) or
      (Q1.FindField ('Y_CREDIT').AsFloat < 0)) or (Q1.FindField ('Y_CREDIT').AsFloat>0)) and (Q1.FindField ('Y_DEBIT').AsFloat=0) then
           TMvt^.Sens := 'C';

      TMvt^.Typeecriture := Q1.FindField ('Y_QUALIFPIECE').asstring;
      TMvt^.NumPiece := IntToStr(Q1.FindField ('Y_NUMEROPIECE').asinteger);
      TMvt^.Devise := Q1.FindField ('Y_DEVISE').asstring;
      TMvt^.TauxDev := AlignDroite(StrfMontant(Q1.FindField ('Y_TAUXDEV').asfloat,20,9,'',False),10);

      OkEuro:=Q1.FindField('Y_DATECOMPTABLE').AsDateTime>=V_PGI.DateDebutEuro ;
      If OkEuro Then
         CodeMontant := 'FDE'
      else
          CodeMontant := 'FD-';
      EnDevise:=RecupDevise(TMvt^.Devise,DecDev,Quotite,TDev) ;
      If EnDevise Then
      BEGIN
           CodeMontant:='DFE' ;
           TMvt^.Montant1:=QuelMontant(Q1,'Y',1,DecDev,TMvt^.Sens) ;
           TMvt^.Montant2:=QuelMontant(Q1,'Y',0,V_PGI.OkDecV,TMvt^.Sens) ;
           TMvt^.Montant3:=QuelMontant(Q1,'Y',2,V_PGI.OkDecE,TMvt^.Sens) ;
           If VH^.TenueEuro Then CodeMontant:='DEF' ;
      END Else
      BEGIN
           If OkEuro Then
           BEGIN
                EnMontantOppose:=Q1.FindField('Y_SAISIEEURO').AsString='X' ;
                If EnMontantOppose Then
                BEGIN
                     If VH^.TenueEuro Then
                          CodeMontant:='FE-'
                     Else
                          CodeMontant:='EF-' ;
                     TMvt^.Montant2:=QuelMontant(Q1,'Y',0,V_PGI.OkDecV,TMvt^.Sens) ;
                     TMvt^.Montant1:=QuelMontant(Q1,'Y',2,V_PGI.OkDecE,TMvt^.Sens) ;
                     TMvt^.Montant3:=Format_String(' ',20) ;
                END Else
                BEGIN
                      If VH^.TenueEuro Then
                           CodeMontant:='EF-'
                      Else
                           CodeMontant:='FE-' ;
                      TMvt^.Montant1:=QuelMontant(Q1,'Y',0,V_PGI.OkDecV,TMvt^.Sens) ;
                      TMvt^.Montant2:=QuelMontant(Q1,'Y',2,V_PGI.OkDecE,TMvt^.Sens) ;
                      TMvt^.Montant3:=Format_String(' ',20) ;
                END ;
           END Else
           BEGIN
               CodeMontant:='FE-' ;
               TMvt^.Montant1:=QuelMontant(Q1,'Y',0,V_PGI.OkDecV,TMvt^.Sens) ;
               TMvt^.Montant2:=QuelMontant(Q1,'Y',2,V_PGI.OkDecE,TMvt^.Sens) ;
               TMvt^.Montant3:=Format_String(' ',20) ;
           END ;
     END ;
   TMvt^.CodeMontant := Codemontant;
   if (EtablissUnique) and (stArg <> '') then
      TMvt^.Etablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', False)
   else
      TMvt^.Etablissement := Q1.FindField ('Y_ETABLISSEMENT').asstring;
   TMvt^.Axe := Q1.FindField ('Y_AXE').asstring;

//   TMvt^.Numeche := IntToStr(Q1.FindField ('Y_NUMECHE').asinteger);
// traitement des devises
  If OkEuro Then
     BEGIN
     SaisieOpposee:=Q1.FindField('Y_SAISIEEURO').AsString='X' ;
     If TMvt^.Devise=V_PGI.DevisePivot Then
        BEGIN
        If ((VH^.TenueEuro) And (Not SaisieOpposee)) Or
           ((Not VH^.TenueEuro) And (SaisieOpposee)) Then TMvt^.Devise:='EUR' ;
        END ;
//     TMvt^.TauxDev:=FloatToStr(Q1.FindField('Y_TAUXDEV').AsFloat/V_PGI.TauxEuro) ;
     END Else
//     TMvt^.TauxDev:=FloatToStr(Q1.FindField('Y_TAUXDEV').AsFloat/V_PGI.TauxEuro) ;
     if FormatEnvoie = 'ETE' then // format étendu
     begin
         TMvt^.RefExterne := Q1.findField('Y_REFEXTERNE').Asstring;
         TMvt^.Daterefexterne := FormatDateTime('ddmmyyyy',Q1.FindField ('Y_DATEREFEXTERNE').AsDateTime);
         TMvt^.Datecreation := FormatDateTime('ddmmyyyy',Q1.FindField ('Y_DATECREATION').AsDateTime);
         TMvt^.Societe := Q1.findField('Y_SOCIETE').Asstring;
         TMvt^.Affaire := Q1.findField('Y_AFFAIRE').Asstring;
         TMvt^.Datetauxdev := FormatDateTime('ddmmyyyy',Q1.FindField ('Y_DATETAUXDEV').AsDateTime);
         TMvt^.Ecranouveau := Q1.findField('Y_ECRANOUVEAU').Asstring;
         TMvt^.Qte1 := AlignDroite(StrfMontant(Q1.FindField('Y_QTE1').AsFloat,20,4,'',False),20);
         TMvt^.Qte2 := AlignDroite(StrfMontant(Q1.FindField('Y_QTE2').AsFloat,20,4,'',False),20);
         TMvt^.QualifQte1 := Q1.findField('Y_QUALIFQTE1').Asstring;
         TMvt^.QualifQte2 := Q1.findField('Y_QUALIFQTE2').Asstring;
         TMvt^.Reflibre := Q1.findField('Y_REFLIBRE').Asstring;
         TMvt^.Contrepartigen := Q1.findField('Y_CONTREPARTIEGEN').Asstring;
         TMvt^.contrepartiaux := Q1.findField('Y_CONTREPARTIEAUX').Asstring;
         TMvt^.LT0 := Q1.findField('Y_LIBRETEXTE0').Asstring;
         TMvt^.LT1 := Q1.findField('Y_LIBRETEXTE1').Asstring;
         TMvt^.LT2 := Q1.findField('Y_LIBRETEXTE2').Asstring;
         TMvt^.LT3 := Q1.findField('Y_LIBRETEXTE3').Asstring;
         TMvt^.LT4 := Q1.findField('Y_LIBRETEXTE4').Asstring;
         TMvt^.LT5 := Q1.findField('Y_LIBRETEXTE5').Asstring;
         TMvt^.LT6 := Q1.findField('Y_LIBRETEXTE6').Asstring;
         TMvt^.LT7 := Q1.findField('Y_LIBRETEXTE7').Asstring;
         TMvt^.LT8 := Q1.findField('Y_LIBRETEXTE8').Asstring;
         TMvt^.LT9 := Q1.findField('Y_LIBRETEXTE9').Asstring;
         TMvt^.TA0 := Q1.findField('Y_TABLE0').Asstring;
         TMvt^.TA1 := Q1.findField('Y_TABLE1').Asstring;
         TMvt^.TA2 := Q1.findField('Y_TABLE2').Asstring;
         TMvt^.TA3 := Q1.findField('Y_TABLE3').Asstring;
         TMvt^.LM0 := Q1.findField('Y_LIBREMONTANT0').Asstring;
         TMvt^.LM1 := Q1.findField('Y_LIBREMONTANT1').Asstring;
         TMvt^.LM2 := Q1.findField('Y_LIBREMONTANT2').Asstring;
         TMvt^.LM3 := Q1.findField('Y_LIBREMONTANT3').Asstring;
         TMvt^.LD :=  FormatDateTime('ddmmyyyy',Q1.FindField ('Y_LIBREDATE').AsDateTime);
         TMvt^.LB0 := Q1.findField('Y_LIBREBOOL0').Asstring;
         TMvt^.LB1 := Q1.findField('Y_LIBREBOOL1').Asstring;
         TMvt^.Conso := Q1.findField('Y_CONSO').Asstring;
     end;
   ListeMouvt.Add(TMvt) ;
   Q1.NExt ;
END ;
Ferme(Q1) ;
if Where <> '' then Wheres := 'Where '+ Where;
//ExecuteSql('UPDATE ANALYTIQ SET Y_EXPORTE="X" '+Wheres) ;

 j:= gettickcount;
// showmessage(inttostr(j-i));
end;


// mode de réglement
procedure TOF_ExportCom.Chargementregle(Listeregle : TList;Where : string);
var
Q1                            :TQuery;
TRegle                        :Preglement;
begin
     Q1 := Opensql ('SELECT * from MODEREGL', TRUE);
     while not Q1.EOF do
     begin
          System.new (TRegle);
          TRegle^.code := Q1.FindField ('MR_MODEREGLE').asstring;
          TRegle^.libelle := Q1.FindField ('MR_LIBELLE').asstring;
          TRegle^.apartirde := Q1.FindField ('MR_APARTIRDE').asstring;
          TRegle^.plusjour := Q1.FindField ('MR_PLUSJOUR').asstring;
          TRegle^.arrondijour := Q1.FindField ('MR_ARRONDIJOUR').asstring;
          TRegle^.nbeche := IntTostr(Q1.FindField ('MR_NOMBREECHEANCE').asinteger);
          TRegle^.separepar := Q1.FindField ('MR_SEPAREPAR').asstring;
          TRegle^.montantmin := Q1.FindField ('MR_MONTANTMIN').asstring;
          TRegle^.remplacemin := Q1.FindField ('MR_REMPLACEMIN').asstring;
          TRegle^.mp1 :=  Q1.FindField ('MR_MP1').asstring;
          TRegle^.taux1 := FloatToStr(Q1.FindField('MR_TAUX1').AsFloat);
          TRegle^.mp2 :=  Q1.FindField ('MR_MP2').asstring;
          TRegle^.taux2 := FloatToStr(Q1.FindField('MR_TAUX2').AsFloat);
          TRegle^.mp3 :=  Q1.FindField ('MR_MP3').asstring;
          TRegle^.taux3 := FloatToStr(Q1.FindField('MR_TAUX3').AsFloat);
          TRegle^.mp4 :=  Q1.FindField ('MR_MP4').asstring;
          TRegle^.taux4 := FloatToStr(Q1.FindField('MR_TAUX4').AsFloat);
          TRegle^.mp5 :=  Q1.FindField ('MR_MP5').asstring;
          TRegle^.taux5 := FloatToStr(Q1.FindField('MR_TAUX5').AsFloat);
          TRegle^.mp6 :=  Q1.FindField ('MR_MP6').asstring;
          TRegle^.taux6 := FloatToStr(Q1.FindField('MR_TAUX6').AsFloat);
          TRegle^.mp7 :=  Q1.FindField ('MR_MP7').asstring;
          TRegle^.taux7 := FloatToStr(Q1.FindField('MR_TAUX7').AsFloat);
          TRegle^.mp8 :=  Q1.FindField ('MR_MP8').asstring;
          TRegle^.taux8 := FloatToStr(Q1.FindField('MR_TAUX8').AsFloat);
          TRegle^.mp9 :=  Q1.FindField ('MR_MP9').asstring;
          TRegle^.taux9 := FloatToStr(Q1.FindField('MR_TAUX9').AsFloat);
          TRegle^.mp10 :=  Q1.FindField ('MR_MP10').asstring;
          TRegle^.taux10 := FloatToStr(Q1.FindField('MR_TAUX10').AsFloat);
          TRegle^.mp11 :=  Q1.FindField ('MR_MP11').asstring;
          TRegle^.taux11 :=FloatToStr(Q1.FindField('MR_TAUX11').AsFloat);
          TRegle^.mp12 :=  Q1.FindField ('MR_MP12').asstring;
          TRegle^.taux12 :=FloatToStr(Q1.FindField('MR_TAUX12').AsFloat);
          Listeregle.add (TRegle);
          Q1.next;
      end;
      ferme (Q1);
end;

procedure ExoToDatesCom ( Exo : String3 ; var ED1,ED2 : string ) ;
Var D1,D2 : TDateTime ;
    Q     : TQuery;
    Okok  : boolean ;
BEGIN
if EXO='' then Exit ;
Okok:=True ; D1:=Date ; D2:=Date ;
If EXO=VH^.Precedent.Code Then BEGIN D1:=VH^.Precedent.Deb ; D2:=VH^.Precedent.Fin ; END Else
If EXO=VH^.EnCours.Code Then BEGIN D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ; END Else
If EXO=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END Else
   BEGIN
   Q:=OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo+'"' ,TRUE) ;
   if Not Q.EOF then
      BEGIN
      D1:=Q.FindField('EX_DATEDEBUT').asDateTime ; D2:=Q.FindField('EX_DATEFIN').asDateTime ;
      END else Okok:=False ;
   Ferme(Q) ;
   END;
if Okok then BEGIN ED1 :=DateToStr(D1) ; ED2 :=DateToStr(D2) ; END ;
END ;

procedure TOF_ExportCom.RemplirEcritureBal (var TMvt : PListeMouvt; Compte, comptecol, nature, libelle : string; var MttSolde : double);
var
CodeMontant                   : string;
Date1,Date2                   : string;
begin
(*     if TMvt^.journal = 'AA' then
     begin
          ExoToDatesCom(Trans.Exo, Date1, Date2);
          TMvt^.Datecomptable := FormatDateTime('ddmmyyyy', StrToDate(Date1));
     end
     else
*)
     TMvt^.Datecomptable := FormatDateTime('ddmmyyyy',Trans.Dateecr2);

     TMvt^.Typepiece := 'OD';
     TMvt^.General := Compte;
     TMvt^.AuxSect := comptecol;

     if (nature = 'COF') and (TMvt^.AuxSect = '') then  TMvt^.AuxSect := GetParamSocSecur ('SO_FOUATTEND', False);
     if (nature = 'COC') and (TMvt^.AuxSect = '') then TMvt^.AuxSect := GetParamSocSecur ('SO_CLIATTEND', False);
     if (nature = 'COD') and  (TMvt^.AuxSect = '') then TMvt^.AuxSect := GetParamSocSecur ('SO_DIVATTEND', False);
     if (nature = 'COS') and (TMvt^.AuxSect = '') then TMvt^.AuxSect := GetParamSocSecur ('SO_SALATTEND', False);

     if TMvt^.AuxSect <> '' then TMvt^.Typecpte:= 'X' else TMvt^.Typecpte:=' ' ;
     TMvt^.Libelle := libelle;
     TMvt^.Echeance := FormatDateTime('ddmmyyyy',iDate1900);
     TMvt^.Typeecriture := 'N';
     TMvt^.Devise := V_PGI.DevisePivot;
     TMvt^.TauxDev := '1';
     if TMvt^.Devise = 'FRF' then   CodeMontant := 'F--';
     if TMvt^.Devise = 'EUR' then   CodeMontant := 'E--';
     TMvt^.CodeMontant := Codemontant;
     TMvt^.Etablissement := Trans.Etabi;
     if TMvt^.Etablissement = '' then
        TMvt^.Etablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', False);
     TMvt^.Axe := '';
     TMvt^.Montant2:=Format_String(' ',20) ;
     TMvt^.Montant3:=Format_String(' ',20) ;
     if (arrondi (MttSolde, V_PGI.OKDECV)) > 0 then TMvt^.Sens := 'D';
     if (arrondi (MttSolde, V_PGI.OKDECV)) < 0 then TMvt^.Sens := 'C';
     if MttSolde < 0.0 then MttSolde := MttSolde * (-1);
     TMvt^.Montant1 := StrfMontant(MttSolde,20,V_PGI.OkDecV,'',False);
end;

function Comparejournal (Item1,Item2:Pointer) : integer;
var
TMvt1,TMvt2             : PListeMouvt;
begin
  if Item1 = nil then Result := -1;
  if Item2 = nil then Result := 1;
  if (Item1 = nil) or (Item2 = nil) then exit;
  Result := 0;

  TMvt1 := Item1; TMvt2 := Item2;
  if TMvt1^.Journal > TMvt2^.Journal then Result := 1
  else if TMvt1^.Journal < TMvt2^.Journal then Result := -1
  else
  if TMvt1^.Journal = TMvt2^.Journal then
  begin
       if TMvt1^.General > TMvt2^.General then Result := 1
       else if TMvt1^.General < TMvt2^.General then Result := -1
       else
       begin
            if TMvt1^.General = TMvt2^.General then
            begin
                      if TMvt1^.AuxSect > TMvt2^.AuxSect then Result := 1
                      else
                      if TMvt1^.AuxSect < TMvt2^.AuxSect then Result := -1;
            end;
       end;
  end;
end;


procedure TOF_ExportCom.Traitement_Ventil (var TMvt : PListeMouvt; var ListeMouvt : TList; Exo,Compte, Collectif, nature, libelle, Jrl : string);
var
MttSolde : double;
SQl      : string;
Q2       : TQuery;
begin
// if (ctxPCL in V_PGI.PGIContexte) then
     Sql := 'SELECT Y_GENERAL, Y_AUXILIAIRE, Y_EXERCICE, ' +
     'sum(Y_DEBIT) as TOTDEBIT, sum(Y_CREDIT) as TOTCREDIT, Y_SECTION, Y_AXE '+
     'FROM ANALYTIQ ' +
     'WHERE  Y_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND Y_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND Y_EXERCICE="'+Exo +   '" AND Y_ECRANOUVEAU="N"  AND Y_GENERAL="' + Compte + '"'+
     ' GROUP BY Y_GENERAL,Y_AUXILIAIRE,Y_EXERCICE, Y_SECTION, Y_AXE '+
     ' ORDER BY Y_GENERAL,Y_AUXILIAIRE,Y_EXERCICE, Y_SECTION, Y_AXE';
(* else
     Sql := 'SELECT Y_GENERAL, Y_AUXILIAIRE, Y_EXERCICE, ' +
     'sum(Y_DEBIT) as TOTDEBIT, sum(Y_CREDIT) as TOTCREDIT, Y_ECRANOUVEAU, Y_SECTION, Y_AXE '+
     'FROM ANALYTIQ ' +
     'WHERE  Y_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND Y_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND Y_EXERCICE="'+Exo +   '" AND Y_ECRANOUVEAU="N"  AND Y_GENERAL="' + Compte + '"'+
     ' GROUP BY Y_GENERAL,Y_AUXILIAIRE,Y_EXERCICE, Y_ECRANOUVEAU,Y_SECTION, Y_AXE '+
     ' ORDER BY Y_GENERAL,Y_AUXILIAIRE,Y_EXERCICE, Y_ECRANOUVEAU,Y_SECTION, Y_AXE';
*)
     MttSolde := 0.0;
     Q2 := OpenSql (Sql, TRUE);
     while not Q2.EOF do
     begin
          MttSolde := arrondi (Q2.FindField ('TOTDEBIT').asFloat - Q2.FindField ('TOTCREDIT').asFloat, V_PGI.OKDECV);
          if MttSolde <> 0.0 then
          begin
            System.New (TMvt);
            (*if not (ctxPCL in V_PGI.PGIContexte) then
            begin
                     if (Q2.FindField ('Y_ECRANOUVEAU').asstring = 'H' )or
                     (Q2.FindField ('Y_ECRANOUVEAU').asstring ='OAN') then
                                   TMvt^.Journal := 'AA'
                     else
                                   TMvt^.Journal := Jrl;
            end
            else
            *)
            TMvt^.Journal := Jrl;
            RemplirEcritureBal (TMvt, Compte, Collectif, nature, libelle, MttSolde);
            TMvt^.AuxSect := Q2.FindField ('Y_SECTION').asstring;
            TMvt^.axe     := Q2.FindField ('Y_AXE').asstring;
            TMvt^.Typecpte:= 'A';
            ListeMouvt.Add(TMvt) ;
          end;
          Q2.next;
     end;
     ferme (Q2);

end;

procedure TOF_ExportCom.ChargementBalance(ListeMouvt : TList; Var F : TextFile);
var
TMvt                          : PListeMouvt;
DecDev                        : integer ;
Jrl                           : string;
MttSolde                      : double;
Compte                        : string;
Exo                           : string;
nature,libelle, Collectif     : string;
OkAuxi                        : Boolean;
Q1                            : TQuery;
SQL                           : string;
EAUXILIAIRE                   : string;
EQUALIFPIECE,Etabliss         : string;
begin

 OkAuxi := TCheckBox (GetControl ('AUXILIARE')).checked;
 Jrl := ReadTokenSt(Trans.Jr);
 if (Jrl = '') or (Jrl = '<<Tous>>')  then
 begin
     Trans.Jr := 'CO';
     Jrl := 'CO';
 end;
 DecDev:=V_PGI.OkDecV ;
 if Trans.Exo <> '' then Exo := Trans.Exo
 else Exo := VH^.Encours.Code;
 if Trans.Dateecr1 = iDate1900 then Trans.Dateecr1 := VH^.Encours.Deb;
 if Trans.Dateecr2 = iDate1900 then Trans.Dateecr2 := VH^.Encours.Fin;
 if Okauxi then  EAUXILIAIRE := 'E_AUXILIAIRE,'
 else EAUXILIAIRE := ' ';

 if TCheckBox (GetControl ('BNORMAL')).checked  then
    EQUALIFPIECE := ' (E_QUALIFPIECE="N"  ' ;
 if TCheckBox (GetControl ('BSIMULE')).checked  then
 begin
    if EQUALIFPIECE = '' then
       EQUALIFPIECE := ' (E_QUALIFPIECE="S"  '
    else
       EQUALIFPIECE := EQUALIFPIECE +  'or E_QUALIFPIECE="S" ';
 end;
 if TCheckBox (GetControl ('BSITUATION')).checked  then
 begin
    if EQUALIFPIECE = '' then
       EQUALIFPIECE := ' (E_QUALIFPIECE="U"  '
    else
       EQUALIFPIECE := EQUALIFPIECE +  'or E_QUALIFPIECE="U" ';
 end;
 if EQUALIFPIECE <> '' then
       EQUALIFPIECE := EQUALIFPIECE + ') ';

 if Trans.Etabi <> '' then Etabliss := 'AND (E_ETABLISSEMENT="'+ Trans.Etabi+'") '
 else Etabliss := '';


 Sql := 'SELECT E_GENERAL,'+ EAUXILIAIRE + 'E_EXERCICE, ' +
     'sum(E_DEBIT) as TOTDEBIT, sum(E_CREDIT) as TOTCREDIT, sum(E_DEBITEURO) as TOTDEBITEUR, sum(E_CREDITEURO) as TOTCREDITEUR, E_QUALIFPIECE,'+
     ' G_NATUREGENE,G_LIBELLE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
     'WHERE  E_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND E_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND E_EXERCICE="'+Exo+'" AND '+ EQUALIFPIECE +
     Etabliss +
     ' GROUP BY E_GENERAL,'+ EAUXILIAIRE+ 'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE '+
     ' ORDER BY E_GENERAL,'+ EAUXILIAIRE+'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE';
(*
     Sql := 'SELECT E_GENERAL,'+ EAUXILIAIRE + 'E_EXERCICE, ' +
     'sum(E_DEBIT) as TOTDEBIT, sum(E_CREDIT) as TOTCREDIT, sum(E_DEBITEURO) as TOTDEBITEUR, sum(E_CREDITEURO) as TOTCREDITEUR, E_QUALIFPIECE,'+
     ' G_NATUREGENE,G_LIBELLE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
     'WHERE  E_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND E_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND E_EXERCICE="'+Exo+'" AND (E_QUALIFPIECE="N" ) ' +
     ' GROUP BY E_GENERAL,'+ EAUXILIAIRE+ 'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE '+
     ' ORDER BY E_GENERAL,'+ EAUXILIAIRE+'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE'
*)
 Q1 := OpenSql (Sql, TRUE);

 MttSolde := 0.0;
 while not Q1.EOF do
 begin
      Compte  := Q1.FindField ('E_GENERAL').asstring;
      if (Compte = GetParamSocSecur ('SO_ECCEUROCREDIT', False)) or  (Compte = GetParamSocSecur ('SO_ECCEURODEBIT', False))  then
         Compte := GetParamSocSecur ('SO_GENATTEND', False);
      Nature  := Q1.FindField ('G_NATUREGENE').asstring;
      Libelle := Q1.FindField ('G_LIBELLE').asstring;
      if Okauxi then
      Collectif  := Q1.FindField ('E_AUXILIAIRE').asstring
      else
      Collectif := '';

      if Collectif <> '' then
         AfficheListeCom('Compte : ' + Collectif, TListBox(GetControl('LISTEEXPORT')))
      else
          AfficheListeCom('Compte : ' + Compte, TListBox(GetControl('LISTEEXPORT')));
      MttSolde := arrondi (Q1.FindField ('TOTDEBIT').asFloat - Q1.FindField ('TOTCREDIT').asFloat,V_PGI.OKDECV);
      if MttSolde <> 0.0 then
      begin
                 System.New (TMvt);
                 (*if not (ctxPCL in V_PGI.PGIContexte) then
                 begin
                     if (Q1.FindField ('E_ECRANOUVEAU').asstring = 'H' )or
                     (Q1.FindField ('E_ECRANOUVEAU').asstring ='OAN') then
                                   TMvt^.Journal := 'AA'
                     else
                                   TMvt^.Journal := Jrl;
                 end
                 else *)
                 TMvt^.Journal := Jrl;
                 RemplirEcritureBal (TMvt, Compte, Collectif, nature, libelle, MttSolde);
                 ListeMouvt.Add(TMvt) ;
                 // pour analytiq
                if (GetControlText('BANA') = 'X') then  // pour ventilation analytique
                begin
                     Traitement_Ventil (TMvt, ListeMouvt,Exo,Compte, Collectif, nature, libelle, Jrl);
                end;

      end;
      Q1.next;
    END ;
    if (ListeMouvt <> nil) and (ListeMouvt.Count > 0 ) then
            ListeMouvt.Sort(CompareJournal);
    Ferme (Q1);
end;


procedure TOF_ExportCom.FileClickzip;
var
  FileName       : String ;
  Commentaire    : String ;
  TheToz         : TOZ;
  Password       : string;
  Filearchive    : string;
begin
  Password := '';
    // Récupération du nom du fichier a insérer
    //
  FileName := GetControlText('FICHENAME');
  Filearchive := ReadTokenPipe (Filename, '.');
  Filearchive := Filearchive + '.zip';

  TheToz := TOZ.Create ;
  try
    if TheToz.OpenZipFile ( Filearchive, moCreate ) then
    begin
        if TheToz.OpenSession ( osAdd ) then
        begin
            if TheToz.ProcessFile ( GetControlText('FICHENAME'), Commentaire ) then
              begin
              TheToz.CloseSession ;
              end
              else
              begin
              HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '' ) ;
              TheToz.CancelSession ;
              end ;
        end
        else
        HShowMessage ( '0;Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '' ) ;
    end
    else
    begin
      HShowMessage ( '0;Erreur;Erreur création du fichier archive : ' + 'archive.zip' + ' impossible;E;O;O;O', '', '' ) ;
      Exit ;
    end ;
 EXCEPT
    On E: Exception do
      begin
      ShowMessage ( 'TozError : ' + E.Message ) ;
      TheToz.Free;
      end ;
 END ;

end;

procedure TOF_ExportCom.SauvegardeQueParametre(var F: TextFile);
var
ListeJalB, ListeCpteaux, ListeSection,ListeRib : TList;
begin
        if (Trans.Pgene = 'X') then
        begin
          AfficheListeCom('Comptes généraux', TListBox(GetControl('LISTEEXPORT')));
          if ListeCpteGen = nil then
          begin
               ListeCpteGen:=TList.Create;
               ChargementCpte(ListeCpteGen,'' );
          end;
          EcritureCpptegen(ListeCpteGen,F);
          ListeCpteGen.Free ;
        end;
        if (Trans.ptiers = 'X') then
        begin
          AfficheListeCom('Comptes Tiers', TListBox(GetControl('LISTEEXPORT')));
          ListeCpteaux:=TList.Create;
          ListeRib:=TList.Create;
          ChargementCpteaux(ListeCpteaux,ListeRib,'');
          EcritureCppteaux(ListeCpteaux,F);
          ListeCpteaux.Free ;
          if TOBTiers <> nil then TOBTiers.free;
          // ajout me 23/05/2003
          EcritureRib(ListeRib,F);
          ListeRib.free;
        end;

       if (TCheckBox (GetControl ('PSECTION')).checked)  then
       begin
           AfficheListeCom('Sections analytiques', TListBox(GetControl('LISTEEXPORT')));
           ListeSection := TList.Create;
           ChargementTableSection(ListeSection, '');
           Ecrituretablesection(ListeSection,F);
           EcritureSectionana(F);
           ListeSection.Free ;
       end;

       if (TCheckBox (GetControl ('PJRL')).checked)  then
       begin
          ListeJalB := TList.Create;
          AlimJalB(ListeJalB,'') ;
          EcritureJournal(ListeJalB, F);
          ListeJalB.Free ;
       end;
end;

procedure TOF_ExportCom.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10    : bClickValide(Sender);
  end;
end;

function TOF_ExportCom.Creationparamcol : Boolean;
var
i,ii      : integer;
Pcol      : ^ar_colreg;
Tcor      : TOB;
Cpte      : string;
begin
     Result := TRUE;
     ExecuteSQL('DELETE from CORRESP Where CR_TYPE="SIS"');
     for i := 1 to FGrille.RowCount - 1 do
//     for i := 1 to FGrilleFou.RowCount - 1 do
     begin
          if Trim(Copy (FGrille.Cells[0, i],3,7)) <> '' then
          begin
            Tcor := TOB.Create('CORRESP', Nil, -1);
            system.New(Pcol);
            Cpte := FGrille.Cells[0, i];
            Pcol^.cpte := AGauche(Cpte,10,'0');
            Cpte := FGrille.Cells[1, i];
            Pcol^.inf := AGauche(Cpte,10,'0');
            Cpte := FGrille.Cells[2, i];
            Pcol^.sup := AGauche(Cpte,10,'Z');
            if Copy (Pcol^.cpte,1,2) = '40' then
               Pcol^.typ := 'F';
            if Copy (Pcol^.cpte,1,2) = '41' then
               Pcol^.typ := 'C';
            Tcor.PutValue('CR_TYPE', 'SIS');
            Tcor.PutValue('CR_CORRESP', Pcol^.cpte);
            Tcor.PutValue('CR_LIBELLE', Pcol^.inf);
            Tcor.PutValue('CR_ABREGE', Pcol^.sup);
            if ExisteSQl ('SELECT * from CORRESP Where CR_TYPE="SIS" and CR_CORRESP="'+ Pcol^.cpte +'"') then
            begin
                 Result := FALSE;
                 Pgiinfo ('Paramétrage incorrect.#10#13Collectif déjà existant : '+Pcol^.cpte,'Export Sisco II');
                 if Pcol <> nil then Dispose (Pcol);    exit;
            end;
            Tcor.InsertDB(nil, TRUE);
            if Pcol <> nil then Dispose (Pcol);
          end;
     end;

end;

function TOF_ExportCom.Existedejacoll (Tg : THGRID; compte : string; var Rr : integer) : Boolean;
var
Atob         : TOB;
ii, indice   : integer;
cpte         : string;
begin
        indice := 0;
        for ii := 1 to Tg.RowCount - 1 do
        begin
          if Trim(Copy (Tg.Cells[0, ii],3,7)) <> '' then
           begin
             cpte := Copy (Tg.Cells[0, ii],0,10);
             Rr := ii;
             if compte = cpte then inc (indice);
             if indice >= 1 then break;
           end;
        end;
        if indice >= 1 then
        begin
             Result := TRUE; exit;
        end;
        Result := FALSE;
end;

function TOF_ExportCom.Existedejaaux (Col,T1,T2 : string) : Boolean;
var
ii                         : integer;
Err                        : Boolean;
Tranche1,Tranche2,cpte     : string;
begin
     Err := FALSE;
     for ii := 1 to FGrille.RowCount - 1 do
     begin
          cpte := Copy (FGrille.Cells[0, ii],0,10);
          if (copy (cpte,1,2)) = (copy (Col,1,2)) then
          begin
                Tranche1 := AGauche(FGrille.Cells[1, ii], 10,'0');
                Tranche2 := AGauche(FGrille.Cells[2, ii], 10,'0');
                if (Tranche1 < T1) then
                begin
                                 if (Tranche2 >= T2) then
                                 begin
                                      Err := TRUE; break;
                                 end;
                end
                else
                begin
                     if (Tranche1 <= T2) then
                     begin
                                  Err := TRUE; break;
                     end;

                end;
          end;
     end;
     if Err then
          PGiBox (' Les tranches de comptes sont déjà définies pour un autre collectif','Paramétrage Sisco II');
     Result := Err;
end;

procedure  TOF_ExportCom.RenWhereEcr( var Where : string; Ext : string);
var
Jourx                   : string;
Datearr                 : TDateTime;
EQUALIFPIECE            : string;
begin
          if Trans.Exo <> '' then
             Where := RendCommandeExo (Ext,Trans.Exo);
          if Trans.Jr <> ''  then
          begin
               Jourx := Trans.Jr;
               if Where <> '' then  Where := Where + ' AND ';
               Where := Where + '('+ RendCommandeJournal(Ext,Jourx)+') ';
          end;
(* A REVOIR pour la prochaine version pour mettre dans paramsoc et gerer e_io avec
          if(THValComboBox(GetControl('NATURETRANSFERT')).Value = 'SYN') then
          begin
               Datearr := StrToDate (GetControlText('DATEECR'));
               if Where <> '' then  Where := Where + ' AND';
               Where := Where + ' ('+Ext+'_DATECOMPTABLE<="'+USDateTime(Datearr)+'")';
          end;
*)
          // ajout me 15-07-2003
         if (THValComboBox(GetControl('NATURETRANSFERT')).Value = 'JRL') then
         begin
                if TCheckBox (GetControl ('BNORMAL')).checked  then
                   EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  ';
                if TCheckBox (GetControl ('BSIMULE')).checked  then
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="S"  '
                   else
                      EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="S" ';
                end;
                if TCheckBox (GetControl ('BSITUATION')).checked  then
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="U"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="U" ';
                end;
                if EQUALIFPIECE <> '' then
                begin
                  EQUALIFPIECE := EQUALIFPIECE + ') ';
                  if Where <> '' then  Where := Where + ' AND ';
                     Where := Where + EQUALIFPIECE;
                end;
          end;
          if Trans.Etabi <> '' then
          begin
               if Where <> '' then  Where := Where + ' AND ';
               Where := Where + '('+Ext+'_ETABLISSEMENT="'+ Trans.Etabi+'") ';
          end;
          if (Trans.Dateecr1 <> iDate1900) and (Trans.Dateecr2 <> iDate1900) then
          begin
               if Where <> '' then  Where := Where + ' AND';
               Where := Where + ' ('+Ext+'_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'"' +
                              ' AND '+Ext+'_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")'
          end
          else
          begin
              if (Trans.Dateecr1 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' ('+Ext+'_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'")';
              end else
              if (Trans.Dateecr2 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' ('+Ext+'_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")';
              end;
          end ;
end;


Initialization
RegisterClasses([TOF_ExportCom]);

end.
