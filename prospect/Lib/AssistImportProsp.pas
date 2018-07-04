unit AssistImportProsp;
//{$UNDEF RECUPCEGID}
interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  Mask, Grids, AssistImport,Hent1,Utob, Hgauge,DBTables, HStatus;

type
  RTypeParametre = record
     code : String ;
     Libelle : String ;
  end;
  TFAsRecupProsp = class(TFAssist)
    Fichier: TTabSheet;
    Grille: TTabSheet;
    PTitre1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    FichierSav: THCritMaskEdit;
    Label3: TLabel;
    OpenFicSav: TOpenDialog;
    PremierProspect: TEdit;
    DernierProspect: TEdit;
    Lborne: TLabel;
    GridPro: THGrid;
    CB_TABPRO: THValComboBox;
    BExtraire: TButton;
    Parametre: TTabSheet;
    ProgressionRecup: TEnhancedGauge;
    Label5: TLabel;
    GrilleRepresentant: TEdit;
    LGrilleInterlocuteur: TLabel;
    GrilleInterlocuteur: TEdit;
    BImportGRC: TButton;
    ProgressionRecup2: TEnhancedGauge;
    Label6: TLabel;
    Label7: TLabel;
    EnhancedGauge1: TEnhancedGauge;
    TypeFiche: TEdit;
    LTypeFiche: TLabel;
    HPreClient: THRadioGroup;
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FichierSavElipsisClick(Sender: TObject);
    procedure BExtraireClick(Sender: TObject);
    procedure BImportGRCClick(Sender: TObject);
  private
    { Déclarations privées }
     Chemin : string;
     passage,recupcegid : boolean;
     NombreLignes : longint;
     NbTActions,NbTPersp :integer;
     TypeAction : Array[1..500] of RTypeParametre ;
     TypePersp : Array[1..500] of RTypeParametre ;
     procedure EclateFichierPgi;
     procedure Grilletechnique;
     function GrilleCode (const CodeProspect,StLigGri: string):string;
     function GrilleListe(const FicProsp:textfile;CodeProspect:string ; var StGri: string):string;
     procedure FormatLaDate(const pos: integer;echeance:boolean;var St: string);
     function FormatLachaine(st : string; long : Integer) : string ;
     function LibelleAction(const code: string): string;
     function LibellePersp(const code: string): string;
     function ProgressionTraitement: longint;
     function InverseDate(const date : string) : string ;
     procedure Correspondance(var stlig : string);
     procedure CiviliteContact(var stlig : string);
     function ExclurePersp (const sttype : string):integer;
     function ExclureAction (const sttype : string):boolean;
     function TraiteCodeClient (const stPro : string):string;
  public
    { Déclarations publiques }
    function  FirstPage : TTabSheet  ; Override ;
    function  NextPage : TTabSheet ; Override ;
  end;
Procedure Assist_ImportProspect ;
Procedure ChangeCodeTiers ;

var
  FAsRecupProsp: TFAsRecupProsp;
  i_NumEcran,NTableCode,NbTablelibre,NbMultilibre : integer;
  st_Chemin: string;

implementation

{$R *.DFM}
Procedure Assist_ImportProspect ;
var
   Fo_Assist : TFAsRecupProsp;
Begin
     Fo_Assist := TFAsRecupProsp.Create (Application);
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

Procedure ChangeCodeTiers ;
var    Q : TQuery;
    CodeAux,CodeTiers,CodeCle :string;
begin
  if PGIAsk('Voulez-vous changer la première lettre pour les tiers clients C par 9 ?','')<>mrYes then exit ;
  Q:=OPENSQL('select t_auxiliaire,t_tiers from tiers where t_auxiliaire like "C%"',true);
  InitMove(Q.recordcount, '');
  while not Q.EOF do
    BEGIN
    CodeCle := Q.Fields[0].AsString;
    CodeAux := Q.Fields[0].AsString;
    CodeTiers := Q.Fields[1].AsString;
    delete (CodeAux,1,1);
    delete (CodeTiers,1,1);
    CodeAux := '9'+CodeAux;
    CodeTiers := '9'+CodeTiers;
    ExecuteSQL('update tiers set t_auxiliaire = '+CodeAux+',t_tiers = '+CodeTiers+' where t_auxiliaire ='+CodeCle+' ');
    Q.Next ;
    MoveCur(false);
    END ;

  FiniMove();
  Ferme(Q);
end;

procedure TFAsRecupProsp.FormShow(Sender: TObject);
begin
  inherited;
    GridPro.ColWidths[0] := 80;
    GridPro.ColWidths[1] := 80;
    GridPro.ColWidths[2] := 140;
    GridPro.ColWidths[3] := 60;
    NbTablelibre:=0;NbMultilibre:=0;
    NombreLignes := 0;
    passage := true;
end;

function TFAsRecupProsp.FirstPage: TTabSheet;
begin
  result := Fichier;
end;

function TFAsRecupProsp.NextPage: TTabSheet;
begin
  if ((getPage = Fichier) and passage) then
  begin
    if  FileExists(FichierSav.text) then
    begin
      Chemin := UpperCase (ExtractFileDir (FichierSav.text)) ;
      screen.cursor:=crHourGlass ;
      Grilletechnique;  // Libellés des grille technique
      if (NombreLignes = 0) then
        NombreLignes := ProgressionTraitement;
      ProgressionRecup.Maxvalue := NombreLignes;
      screen.cursor:=crDefault ;
      result := Grille ;
      passage := false;
    end else
       result := Fichier
  end else
      if P.ActivePage.PageIndex<P.PageCount-1 then result:=P.Pages[P.ActivePage.PageIndex+1] else result:=nil ;
end;

procedure TFAsRecupProsp.bFinClick(Sender: TObject);
var
   Fo_Assist : TFAssistimport;
   begin
    inherited;
    if getPage = Parametre then
       begin
       bFin.Enabled := False;
       screen.cursor:=crHourGlass ;
       EclateFichierPgi;
       screen.cursor:=crDefault ;
       Fo_Assist := TFAssistimport.Create (Application);
  //   Fo_Assist.CMContexte.Text := 'PROSPECT';
       Try
           Fo_Assist.ShowModal;
       Finally
           Fo_Assist.free;
       End;
      //Close ;
    end;
end;

procedure TFAsRecupProsp.EclateFichierPgi;
var
    StPro,StLig,StPgi,StProj,StLibelle,StAffaire,StHisto,StRep : String;
    StBLoc,StPgiB,datesaisie,NaturePersp  : String;
    CodeProspect,CodeClient,CodeIntervenant,Liste: String;
    FLigProsp,FLigFicSav,FLigFicBloc,FPgiComPersp : textfile;
    FPgiProsp,FPgiGri,FPgiContact,FPgiAction,FPgiBloc,FPgiPersp,FPgiPerspHis: textfile;
    FPgiCode,FPgiCom,FPgiProj,FPgiRep,FPgiComAction,FPgiRes,FPgiParActions : textfile;
    Compteur,comptproj,comptIndice,i_ind,i_indL,i_rep1,i_rep2,NumAffaire,NumNature : integer;
    Principale : Char;
    AffProjet,AffIndice : array [0..100, 1..50] of integer;
begin
    StPro := '';
    ProgressionRecup.Progress := 0;
    ProgressionRecup.Visible := true;
//    ProgressionRecup.Maxvalue  := ProgressionTraitement ;

    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);

    AssignFile(FLigFicSav, FichierSav.text);
    Reset (FLigFicSav);

    AssignFile(FPgiProsp, Chemin + '\PROTIERS.pgi');
    Rewrite (FPgiProsp);

    AssignFile(FPgiBloc, Chemin + '\PROBLOC.pgi');
    Rewrite (FPgiBloc);

    AssignFile(FPgiGri, Chemin + '\PROSPECTS.pgi');
    Rewrite (FPgiGri);

    AssignFile(FPgiContact, Chemin + '\PROCONTACT.pgi');
    Rewrite (FPgiContact);

    AssignFile(FPgiAction, Chemin + '\PROACTIONS.pgi');
    Rewrite (FPgiAction);

    AssignFile(FPgiPersp, Chemin + '\PERSPECTIVES.pgi');
    Rewrite (FPgiPersp);

    AssignFile(FPgiPerspHis, Chemin + '\PERSPHISTO.pgi');
    Rewrite (FPgiPerspHis);

    AssignFile(FPgiCom, Chemin + '\PROCOMMUN.pgi');
    Rewrite (FPgiCom);

    AssignFile(FPgiCode, Chemin + '\PROCHOIXCODE.pgi');
    Rewrite (FPgiCode);

    AssignFile(FPgiProj, Chemin + '\PROPROJETS.pgi');
    Rewrite (FPgiProj);

    AssignFile(FPgiRep, Chemin + '\PROCOMMERCIAL.pgi');
    Rewrite (FPgiRep);

    AssignFile(FPgiRes, Chemin + '\RESSOURCE.pgi');
    Rewrite (FPgiRes);

    AssignFile(FPgiComAction, Chemin + '\PROCOMACTIONS.pgi');
    Rewrite (FPgiComAction);

    AssignFile(FPgiComPersp, Chemin + '\PROCOMPERSP.pgi');
    Rewrite (FPgiComPersp);

    AssignFile(FPgiParActions, Chemin + '\PROPARACTIONS.pgi');
    Rewrite (FPgiParActions);

    readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3)<>'BLO') ) do
      readln (FLigFicSav,StLig);

    while (Copy(StPro,1,3) = 'PRO') do
    begin
      CodeProspect :=  Copy(StPro,4,10);
      CodeClient := TraiteCodeClient(StPro);
      { CodeClient := Copy(StPro,14,10);
      if (( (StPro[34]='C') or (StPro[34]='E') ) and (trim(CodeClient)<>'') ) then
        CodeClient := '9'+Copy(StPro,14,9) else CodeClient := '9'+Copy(StPro,4,9);    }
      if (Copy(StLig,1,3) = 'BLO') then
      begin
        while ( not EOF(FLigFicSav) and (Copy(StLig,1,13) < 'BLO'+CodeProspect) ) do
          readln (FLigFicSav,StLig);
        while (not EOF(FLigFicSav) and (Copy(StLig,1,13) = 'BLO'+CodeProspect)) do
        begin
          StPgi := CodeClient;
          StPgi := StPgi + Copy(StLig,13,4);
          StPgi := StPgi + Copy(StLig,19,Length(StLig)-19);
          WriteLn (FPgiBloc,StPgi) ;
          readln (FLigFicSav,StLig);
          ProgressionRecup.Progress := ProgressionRecup.Progress+1;
        end;
      end;
      readln (FLigProsp,StPro);
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) <> 'INL') and (Copy(StLig,1,6) < 'TABACC')) do
      readln (FLigFicSav,StLig);
    while (not EOF(FLigFicSav) and (Copy(StLig,1,6) = 'TABACC') ) do
    begin
      StPgi := Copy(StLig,7,3) + Copy(StLig,17,20);
      WriteLn (FPgiParActions,StPgi) ;
      readln (FLigFicSav,StLig);
      ProgressionRecup.Progress := ProgressionRecup.Progress+1;
    end;

    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) <> 'INL') and (Copy(StLig,1,5) < 'TABGR') ) do
      readln (FLigFicSav,StLig);
    i_indL := -1;
    while (Copy(StLig,1,5) = 'TABGR') do
    begin
      i_ind := Ord(StLig[6])- ord('A');
      if (Copy(GridPro.Cells[1, i_ind],1,5) = 'Liste') then
      begin
         StPgi := 'RM'+ GridPro.Cells[3, i_ind];
         if (i_ind <> i_indL) then
         begin
            StLibelle := 'RLZ' + 'ML' + GridPro.Cells[3, i_ind] + GridPro.Cells[2, i_ind];
            WriteLn (FPgiCode,StLibelle) ;
            ProgressionRecup.Progress := ProgressionRecup.Progress+1;
         end;
      end else
      begin
         StPgi := 'RL'+ GridPro.Cells[3, i_ind];
         if (i_ind <> i_indL) then
         begin
            StLibelle := 'RLZ' + 'CL' + GridPro.Cells[3, i_ind] + GridPro.Cells[2, i_ind];
            WriteLn (FPgiCode,StLibelle) ;
            ProgressionRecup.Progress := ProgressionRecup.Progress+1;
         end;
         if ((GrilleRepresentant.text <> '') and (StLig[6] = GrilleRepresentant.text)) then
         begin
            StRep := 'REP' +  Copy(StLig,7,3) + '001' + Copy(StLig,17,20) ;
            WriteLn (FPgiRep,StRep) ;
            ProgressionRecup.Progress := ProgressionRecup.Progress+1;
         end;
      end;
      StPgi := StPgi + Copy(StLig,7,3);
      StPgi := StPgi + Copy(StLig,17,20);
      WriteLn (FPgiCode,StPgi) ;
      readln (FLigFicSav,StLig);
      ProgressionRecup.Progress := ProgressionRecup.Progress+1;
      i_indL := i_ind ;
    end;

    if ( GrilleInterlocuteur.text <> '' ) then
    begin
      while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) <> 'INL') and (Copy(StLig,1,6) < 'TABIN'+GrilleInterlocuteur.text) ) do
        readln (FLigFicSav,StLig);
      while (not EOF(FLigFicSav) and (Copy(StLig,1,6) = 'TABIN'+GrilleInterlocuteur.text) ) do
      begin
        StPgi := 'FON' + Copy(StLig,7,3) + Copy(StLig,17,20);
        WriteLn (FPgiCode,StPgi) ;
        readln (FLigFicSav,StLig);
        ProgressionRecup.Progress := ProgressionRecup.Progress+1;
      end;
    end;

    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) <> 'INL') and (Copy(StLig,1,6) < 'TABIVT') ) do
      readln (FLigFicSav,StLig);
    while (not EOF(FLigFicSav) and (Copy(StLig,1,6) = 'TABIVT') ) do
    begin
      StRep := 'SAL' +  Copy(StLig,7,3) + Copy(StLig,17,20) ;
      WriteLn (FPgiRes,StRep) ;
      readln (FLigFicSav,StLig);
      ProgressionRecup.Progress := ProgressionRecup.Progress+1;
    end;

    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) <> 'INL') and (Copy(StLig,1,6) < 'TABPER') ) do
      readln (FLigFicSav,StLig);
    while (not EOF(FLigFicSav) and (Copy(StLig,1,6) = 'TABPER') ) do
    begin
      StPgi := 'RTP' + Copy(StLig,7,3) + Copy(StLig,17,20);
      WriteLn (FPgiCode,StPgi) ;
      readln (FLigFicSav,StLig);
      ProgressionRecup.Progress := ProgressionRecup.Progress+1;
    end;

    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) <> 'INL') and (Copy(StLig,1,6) < 'TABSIT') ) do
      readln (FLigFicSav,StLig);
    while (not EOF(FLigFicSav) and (Copy(StLig,1,6) = 'TABSIT') ) do
    begin
      StPgi := 'RSP' + Copy(StLig,7,3) + Copy(StLig,17,20);
      WriteLn (FPgiCom,StPgi) ;
      readln (FLigFicSav,StLig);
      ProgressionRecup.Progress := ProgressionRecup.Progress+1;
    end;


    CloseFile(FLigProsp);
    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);
    readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3)<>'INL') ) do
      readln (FLigFicSav,StLig);

    while (Copy(StPro,1,3) = 'PRO') do
    begin
      CodeProspect :=  Copy(StPro,4,10);
      CodeClient := TraiteCodeClient(StPro);
      {CodeClient := Copy(StPro,14,10);
      if (( (StPro[34]='C') or (StPro[34]='E') ) and (trim(CodeClient)<>'') ) then
        CodeClient := '9'+Copy(StPro,14,9) else CodeClient := '9'+Copy(StPro,4,9);  }

      if (Copy(StLig,1,3) = 'INL') then
      begin
        while ( not EOF(FLigFicSav) and (Copy(StLig,1,13) < 'INL'+CodeProspect) ) do
          readln (FLigFicSav,StLig);
        Principale := 'X';
        while (not EOF(FLigFicSav) and (Copy(StLig,1,13) = 'INL'+CodeProspect)) do
        begin
          StPgi := CodeClient + Copy(StLig,14,3);    //Code prospect
          StPgi := StPgi + StPro[34];  // nature auxiliaire
          StPgi := StPgi + Principale;
          Principale := '-';
          StPgi := StPgi + Copy(StLig,17,Length(StLig)-17);
          WriteLn (FPgiContact,StPgi) ;
          readln (FLigFicSav,StLig);
          ProgressionRecup.Progress := ProgressionRecup.Progress+1;
        end;
      end;
      readln (FLigProsp,StPro);
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    CloseFile(FLigProsp);
    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);
    readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3)<>'GRI') and (Copy(StLig,1,3)<>'ACT')) do
      readln (FLigFicSav,StLig);

    while (Copy(StPro,1,3) = 'PRO') do
    begin
      CodeProspect :=  Copy(StPro,4,10);
      CodeClient := TraiteCodeClient(StPro);
     {CodeClient := Copy(StPro,14,10);  15-11-2001
      if (( (StPro[34]='C') or (StPro[34]='E') ) and (trim(CodeClient)<>'') ) then
        CodeClient := '9'+Copy(StPro,14,9) else CodeClient := '9'+Copy(StPro,4,9); }

      if (StPro[297]='O') then StPro[297]:='X' else StPro[297]:='-';
      StPgi := Copy(StPro,614,78);
      FormatLaDate(315,False,StPro);
      FormatLaDate(323,False,StPro);
      FormatLaDate(331,False,StPro);

      Liste := CodeClient + Copy(StPro,339,275);
      if (trim (StPgi)<>'') then
      begin
        while ( not EOF(FLigFicSav) and (Copy(StLig,1,3) = 'GRI') and (Copy(StLig,1,13) < 'GRI'+CodeProspect) ) do
          readln (FLigFicSav,StLig);
        Liste := Liste + GrilleCode(Copy(StPro,4,10), StPgi);
        Liste := Liste + GrilleListe(FLigFicSav, CodeProspect, StLig);
        WriteLn (FPgiGri,Liste) ;
        ProgressionRecup.Progress := ProgressionRecup.Progress+1;
      end;
      WriteLn (FPgiProsp,CodeClient+CodeProspect+Copy(StPro,14,Length(StPro)-14)) ;
      readln (FLigProsp,StPro);
      ProgressionRecup.Progress := ProgressionRecup.progress + 1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;


    CloseFile(FLigProsp);
    AssignFile(FLigProsp, FichierSav.text);     // Prospect
    Reset (FLigProsp);
    readln (FLigProsp,StPro);
    AssignFile(FLigFicBloc, FichierSav.text);   //Commentaire Action
    Reset (FLigFicBloc);
    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3)<>'ACT') ) do
      readln (FLigFicSav,StLig);
    while ( not EOF(FLigFicBloc) and (Copy(StBLoc,1,3)<>'ACC') ) do
      readln (FLigFicBloc,StBLoc);

   while (Copy(StPro,1,3) = 'PRO') do
    begin
      CodeProspect :=  Copy(StPro,4,10);
      CodeClient := TraiteCodeClient(StPro);
     {CodeClient := Copy(StPro,14,10);
      if (( (StPro[34]='C') or (StPro[34]='E') ) and (trim(CodeClient)<>'') ) then
        CodeClient := '9'+Copy(StPro,14,9) else CodeClient := '9'+Copy(StPro,4,9);  }
      Compteur := 0;
      if (Copy(StLig,1,3) = 'ACT') then
      begin
        while ( not EOF(FLigFicSav) and (Copy(StLig,1,13) < 'ACT'+CodeProspect) ) do
          readln (FLigFicSav,StLig);
        while ( not EOF(FLigFicBloc) and (Copy(StBloc,1,13) < 'ACC'+CodeProspect) ) do
          readln (FLigFicBloc,StBloc);
        while (not EOF(FLigFicSav) and (Copy(StLig,1,13) = 'ACT'+CodeProspect)) do
        begin
          StRep :=  'ACC' + Copy(StLig,4,22);
          inc (Compteur);
          StPgi := Format('%.3u', [Compteur]);                       //NUMACTION
          StPgi := StPgi + LibelleAction('ACC'+Copy(StLig,42,3));    //LIBELLE
          StPgi := StPgi + CodeClient;                               //AUXILIAIRE
          StPgi := StPgi + Copy(StLig,42,3);                         //TYPEACTION
          StPgi := StPgi + Copy(StLig,45,3);                         //INTERVENANT
          if (StLig[14] = 'P') then
          begin
            FormatLaDate(34,true,StLig);
            StPgi := StPgi + Copy(StLig,34,8);
            StPgi := StPgi + 'PRE';                               //ETATACTION Prévu
          end else
          begin
            FormatLaDate(26,true,StLig);
            StPgi := StPgi + Copy(StLig,26,8);
            StPgi := StPgi + 'REA';                               //ETATACTION  Réalisé
          end;
          StPgiB := StPgi + '000' + Copy(StLig,48,60);                     //COMMENTAIRE
          WriteLn (FPgiComAction,StPgiB) ;
          while ( not EOF(FLigFicBloc) and (Copy(StBloc,1,25) = StRep) ) do
          begin
            StPgiB := StPgi +  Copy(StBloc,26,Length(StBloc)-26);                    //COMMENTAIRE
            WriteLn (FPgiComAction,StPgiB) ;
            readln (FLigFicBloc,StBloc);
            ProgressionRecup.Progress := ProgressionRecup.Progress+1;
          end;
          WriteLn (FPgiAction,StPgi) ;

          readln (FLigFicSav,StLig);
          ProgressionRecup.Progress := ProgressionRecup.Progress+1;
        end;
      end;
      readln (FLigProsp,StPro);
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;
    CloseFile(FLigFicBloc);


    CloseFile(FLigProsp);
    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);
    readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigFicSav) and (Copy(StLig,1,3)<>'PER') ) do
      readln (FLigFicSav,StLig);
    compteur := 1; comptproj := 1; StProj := '';

    while ((Copy(StPro,1,3) = 'PRO') and (Copy(StLig,1,3) = 'PER')) do
    begin
      CodeProspect :=  Copy(StPro,4,10);
      CodeClient := TraiteCodeClient(StPro);
     {CodeClient := Copy(StPro,14,10);
      if (( (StPro[34]='C') or (StPro[34]='E') ) and (trim(CodeClient)<>'') ) then
        CodeClient := '9'+Copy(StPro,14,9) else CodeClient := '9'+Copy(StPro,4,9);   }
      StAffaire := '  ';
      comptIndice := 1;
      CodeIntervenant := '   ';
      NaturePersp := '   ';
      if (GrilleRepresentant.text <> '') then
      begin
        i_rep1 := 614 + ( 3 * (ord(GrilleRepresentant.text[1]) - ord('A')) );
        CodeIntervenant := Copy(StPro,i_rep1,3);
      end;

      while ( not EOF(FLigFicSav) and (Copy(StLig,1,13) < 'PER'+CodeProspect) ) do
        readln (FLigFicSav,StLig);
      while (not EOF(FLigFicSav) and (Copy(StLig,1,13) = 'PER'+CodeProspect)) do
      begin
          StPgi := Format('%.6u', [compteur]);//Numéro
          StPgi := StPgi + Copy(StLig,25,3);                             //Nature
          Stlibelle := LibellePersp('PER'+Copy(StLig,25,3));
          StPgi := StPgi + Stlibelle;         //Libellé
          StPgi := StPgi + CodeClient;  //codePro ou code_cli -> AUXILIAIRE
          StPgi := StPgi + StPro[34];         //cli_pro    -> TYPETIERS
          StPgi := StPgi + CodeIntervenant;   //INTERVENANT
          StPgi := StPgi + Format('%-6u', [compteur]); // Numero projet
          FormatLaDate(28,true,StLig);
          FormatLaDate(87,false,StLig);
          StPgi := StPgi + Copy(StLig,28,8);  //dateReal   -> DATEREALISE
          StPgi := StPgi + StLig[46] + '  ';         //situation  -> ETATPER
          StPgi := StPgi + Copy(StLig,36,7);  //caPrevu    -> MONTANTPER
          StPgi := StPgi + Copy(StLig,43,3);  //pcReal     -> POURCENTAGE
          datesaisie := InverseDate(Copy(StLig,14,8));  //dateSaisie -> DATECREATION
          StPgi := StPgi + datesaisie;
          StPgi := StPgi + '31122099';        //Date fin de vie
          if  (trim(Copy(StLig,47,40)) <> '') then
          begin
            StPgiB := Format('%.6u', [compteur]) + Copy(StLig,47,40);// Commentaire
            WriteLn (FPgiComPersp,StPgiB);
          end;
          WriteLn (FPgiPersp,StPgi);
          inc (compteur);
          readln (FLigFicSav,StLig);
          ProgressionRecup.Progress := ProgressionRecup.Progress+1;
      end;     

      readln (FLigProsp,StPro);
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    CloseFile(FLigProsp);
    CloseFile(FLigFicSav);
    CloseFile(FPgiBloc);
    CloseFile(FPgiProsp);
    CloseFile(FPgiGri);
    CloseFile(FPgiContact);
    CloseFile(FPgiAction);
    CloseFile(FPgiPersp);
    CloseFile(FPgiCode);
    CloseFile(FPgiCom);
    CloseFile(FPgiProj);
    CloseFile(FPgiPerspHis);
    CloseFile(FPgiRep);
    CloseFile(FPgiComAction);
    CloseFile(FPgiComPersp);
    CloseFile(FPgiRes);
    CloseFile(FPgiParActions);
end;

function TFAsRecupProsp.GrilleCode (const CodeProspect,StLigGri: string):string;
var
  StPgi,code : string;
  i_gri : integer;
begin
  Result := '';
  StPgi := '';
  for i_gri := 0 to GridPro.RowCount - 1 do
      GridPro.Cells[3, i_gri] := '';

  for i_gri := 0 to 25 do
  begin
    code := Copy(StLigGri,(i_gri*3+1),3);
    if (code <> '+->') then
       GridPro.Cells[3, i_gri] := Code;
  end;
  for i_gri := 0 to 25 do
  begin
    if ( (trim(GridPro.Cells[0, i_gri]) <> '') and (Copy(GridPro.Cells[1, i_gri],1,5) <> 'Liste') ) then
      StPgi := StPgi + GridPro.Cells[3, i_gri];
  end;
  StPgi := FormatLachaine (StPgi,78);
  Result := StPgi;
end;

function TFAsRecupProsp.GrilleListe(const FicProsp:textfile;CodeProspect:string ; var StGri: string):string;
var
  StPgi: string;
  i_gri : integer;
begin

  Result := ''; StPgi := '';
  while (not EOF(FicProsp) and (Copy(StGri,1,3) = 'GRI') and (Copy(StGri,1,13) = 'GRI'+CodeProspect)) do
  begin
    i_gri := Ord(StGri[16]) - ord('A');
    if (trim(Copy(StGri,17,3)) <> '') then
      GridPro.Cells[3, i_gri] := GridPro.Cells[3, i_gri] + Copy(StGri,17,3) + ';';
    readln (FicProsp,StGri);
    ProgressionRecup.Progress := ProgressionRecup.Progress+1;
  end;

  for i_gri := 0 to 25 do
  begin
    if ( (trim(GridPro.Cells[0, i_gri]) <> '') and (Copy(GridPro.Cells[1, i_gri],1,5) = 'Liste') ) then
      StPgi := StPgi + Format_String(GridPro.Cells[3, i_gri],80);
  end;
  Result := StPgi;
end;

procedure TFAsRecupProsp.Grilletechnique;
var
  FLigProsp: textfile;
  St : string;
  i_ind1,i_ind2: integer;
  DerniereTable,FinDeTable : boolean;
begin
  AssignFile(FLigProsp, FichierSav.text);
  Reset (FLigProsp);
  readln (FLigProsp,St);
  NbMultilibre:=0;     NbTablelibre := 0;
  while ( not EOF(FLigProsp) and (Copy(st,1,3)<>'TAB') ) do
      readln (FLigProsp,st);

  i_ind1 := 28;NbTActions:=0;NbTPersp:=0 ; DerniereTable:= False; FinDeTable := False;
  while ( not EOF(FLigProsp) and not FinDeTable ) do
  Begin
    if (Copy(St,1,6) = 'TABACC') then
    begin
      inc(NbTActions);
      TypeAction[NbTActions].code := Copy(St,4,6);
      TypeAction[NbTActions].Libelle := Copy(St,17,20);
    end
    else  if (Copy(St,1,7) = 'TABLGRL') then
    begin
      i_ind2 := Ord(St[9])- ord('A');
      if (trim(Copy(St,17,20)) <> '') then
      begin
        GridPro.Cells[0, i_ind2] := 'Prospect';
        if (St[37]='L') then
        begin
          GridPro.Cells[1, i_ind2] := 'Liste '+St[9];
          GridPro.Cells[3, i_ind2] := format ('%.1u',[NbMultilibre]);
          inc (NbMultilibre);
        end else
        begin
          if (NbTablelibre > 9) then
          begin
              GridPro.Cells[3, i_ind2] := chr (NbTablelibre - 10 + ord('A'));
          end else
              GridPro.Cells[3, i_ind2] := format ('%.1u',[NbTablelibre]);
          inc (NbTablelibre);
          GridPro.Cells[1, i_ind2] := 'Grille '+St[9];
        end;
        GridPro.Cells[2, i_ind2] := Copy(St,17,20);
      end;
    end
    else if (Copy(St,1,7) = 'TABLINL') then
    begin
      GridPro.Cells[0, i_ind1] := 'Intervenant';
      GridPro.Cells[1, i_ind1] := 'INL'+ St[9];
      GridPro.Cells[2, i_ind1] := Copy(St,17,20);
      inc(i_ind1);
    end
    else if (Copy(St,1,6) = 'TABPER') then
    begin
      inc(NbTPersp);
      TypePersp[NbTPersp].code := Copy(St,4,6);
      TypePersp[NbTPersp].Libelle := Copy(St,17,20);
    end
    else if (Copy(St,1,6) = 'TABSIT') then
    begin
      GridPro.Cells[0, i_ind1] := 'Situation';
      GridPro.Cells[1, i_ind1] := Copy(St,7,3);
      GridPro.Cells[2, i_ind1] := Copy(St,17,20);
      inc(i_ind1);
      DerniereTable := true;
    end
    else if (DerniereTable) then FinDeTable := true;
    readln (FLigProsp,st);
  end;
  NTableCode := i_ind1;
  CloseFile(FLigProsp);
end;

function TFAsRecupProsp.LibelleAction (const code: string): string;
var i_par : integer;
begin
  result := '';
  for i_par := 1 to NbTActions do
    if ( TypeAction[i_par].code = code ) then  result := TypeAction[i_par].Libelle;
  result := FormatLachaine (result,20);
end;

function TFAsRecupProsp.LibellePersp(const code: string): string;
var i_par : integer;
begin
  result := '';
  for i_par := 1 to NbTPersp do
    if ( TypePersp[i_par].code = code ) then  result := TypePersp[i_par].Libelle;
  result := FormatLachaine (result,20);
end;

procedure TFAsRecupProsp.FichierSavElipsisClick(Sender: TObject);
begin
  inherited;
  if OpenFicSav.execute then FichierSav.text:=OpenFicSav.FileName;
end;

procedure TFAsRecupProsp.FormatLaDate(const pos: integer;echeance:boolean;var St: string);
var
  StDate1,StDate2 : String;
begin
  StDate1 := Copy(St,pos,8);  //19990327   -> 27031999
  if (echeance) then StDate2 := '31122099' else StDate2 := '01011900';
  if (trim(StDate1) <> '') then
    StDate2 := Copy(StDate1,7,2)+Copy(StDate1,5,2)+Copy(StDate1,1,4) ;
  Delete(St,pos,8);
  Insert(StDate2,St,pos) ;
end;

function TFAsRecupProsp.FormatLachaine(st: string; long: Integer): string;
var
  St1 : String ;
begin
  st1:=Trim(st) ;
  if ((long>0) and (long<Length(St1))) then St1:=Copy(St1,1,long) ;
  While Length(St1)<long Do St1:=St1+' ' ;
    result:=st1 ;
end;

function TFAsRecupProsp.ProgressionTraitement: longint;
var  FLigProsp :textfile;
     st : string;
     long_persp : integer;
     nb_lig : longint;
begin
  AssignFile(FLigProsp, FichierSav.text);
  Reset (FLigProsp);
  nb_lig := 0;long_persp:=0;
  while ( not EOF(FLigProsp)) do
  begin
    readln (FLigProsp,st);
    inc(nb_lig);
  end;
  CloseFile(FLigProsp);
  result := nb_lig;
end;

function TFAsRecupProsp.InverseDate(const date: string): string;
var Stdate : string;
    i_car : integer;
begin
  result := '        ';
  Stdate:= '';
  if (trim(date) <> '') then
  begin
    for i_car := 1 to 8 do
        Stdate := Stdate + chr( ord('9') - ord(date[i_car]) + ord('0') );
    result := Copy(StDate,7,2)+Copy(StDate,5,2)+Copy(StDate,1,4) ;
  end;
end;

function TFAsRecupProsp.TraiteCodeClient(const StPro: string): string;
var CodeClient : string;
begin
    CodeClient := Copy(StPro,14,10);
    if ( (trim(CodeClient)<>'') and (trim(TypeFiche.text) <> '') and (Pos (StPro[34],TypeFiche.text) > 0) )
    then  CodeClient := HPreClient.value+Copy(StPro,14,9)
    else  CodeClient := HPreClient.value+Copy(StPro,4,9);

    result := CodeClient;
end;

procedure TFAsRecupProsp.BExtraireClick(Sender: TObject);
var StPro :string;
    FLigProsp,FLigFicSav : textfile;
    compteur :integer;
    chemin,Civilite: String ;
    i_ind1 : Integer ;
    ListeJuri: TStringList;
begin
//    if ( (getPage <> Fichier) or ((PremierProspect.text='') and (DernierProspect.text='')) )
//      then  exit;

{  AssignFile(FLigProsp, FichierSav.text);
  Reset (FLigProsp);
  Chemin := UpperCase (ExtractFileDir (FichierSav.text)) ;
  AssignFile(FLigFicSav, Chemin + '\Civilite.sav');
  Rewrite (FLigFicSav);
  while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'INL') ) do
    readln (FLigProsp,StPro);
  ListeJuri := TStringList.Create;
  ListeJuri.Sorted := True;

  while ( not EOF(FLigProsp)) and (Copy(StPro,1,3) = 'INL') do
  begin
      Civilite := Copy(StPro,17,5);
      ListeJuri.Add(trim(Civilite));
      readln (FLigProsp,StPro);
  end;
  ListeJuri.sort;
  for i_ind1 := 1 to ListeJuri.Count-1 do
    WriteLn (FLigFicSav,ListeJuri.Strings[i_ind1]);

  ListeJuri.free;
  CloseFile(FLigProsp);
  CloseFile(FLigFicSav);
  exit; }

    bSuivant.Enabled := False;
    ProgressionRecup2.Visible := true;
    ProgressionRecup2.Progress := 5;
    screen.cursor:=crHourGlass ;
    if (NombreLignes = 0) then
      NombreLignes := ProgressionTraitement;
    ProgressionRecup2.Maxvalue := NombreLignes;

    Chemin := UpperCase (ExtractFileDir (FichierSav.text)) ;
    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);

    {
    AssignFile(FLigFicSav, Chemin + '\Action.sav');
    Rewrite (FLigFicSav);
    readln (FLigProsp,StPro);
    compteur := 0;
    while (Copy(StPro,1,3)<>'ACT') do
      readln (FLigProsp,StPro);
    while  not EOF(FLigProsp) and (Copy(StPro,1,3) = 'ACT')  do
    begin
        if ((StPro[14] = 'P') and (Copy(StPro,34,4)=trim(PremierProspect.text)))
        or ((StPro[14] = 'R') and (Copy(StPro,26,4)=trim(PremierProspect.text))) then
        begin
          inc(compteur);
          writeln (FLigFicSav,StPro);
        end;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;
    DernierProspect.text := format ( '%.8u',[compteur]);
    BExtraire.Enabled := False;
    ProgressionRecup2.Visible := false;
    bSuivant.Enabled := True;
    screen.cursor:=crDefault ;

    CloseFile(FLigProsp);
    CloseFile(FLigFicSav);
    exit;   }


    AssignFile(FLigFicSav, Chemin + '\Prospect_e.sav');
    Rewrite (FLigFicSav);

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);

    while (Copy(StPro,1,3) = 'PRO') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'BLO') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);

    while (Copy(StPro,1,3) = 'BLO') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'TAB') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'TAB') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'INL') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'INL') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'GRI') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'GRI') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'ACT') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'ACT') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'ACC') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'ACC') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PER') ) do
      readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (trim(Copy(StPro,4,10)) < trim(PremierProspect.text)) ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'PER') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
      if ((trim(DernierProspect.text) <> '') and (trim(Copy(StPro,4,10)) > DernierProspect.text)) then
      break;
    end;
    BExtraire.Enabled := False;
    ProgressionRecup2.Visible := false;
    bSuivant.Enabled := True;
    screen.cursor:=crDefault ;
    FichierSav.text := Chemin + '\Prospect_e.sav';
    PremierProspect.text:='';DernierProspect.text := '';
    CloseFile(FLigProsp);
    CloseFile(FLigFicSav);
end;


procedure TFAsRecupProsp.BImportGRCClick(Sender: TObject);
var StPro,CodeProsp :string;
    FLigProsp,FLigFicSav ,FDoubleProsp: textfile;
    ListeDoubles: TStringList;
    Index: Integer;
    Fo_Assist : TFAssistimport;
begin
    if getPage <> Fichier then  exit
    else
       begin
       bFin.Enabled := False;
       screen.cursor:=crDefault ;
       Fo_Assist := TFAssistimport.Create (Application);
       Fo_Assist.CMContexte.Text := 'PROSPECT';
       Try
           Fo_Assist.ShowModal;
       Finally
           Fo_Assist.free;
       End;
      Close ;
    end;
    exit;


    bSuivant.Enabled := False;
    Chemin := UpperCase (ExtractFileDir (FichierSav.text)) ;
    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);

    AssignFile(FDoubleProsp, Chemin + '\Tiers_Doubles.log');
    Reset (FDoubleProsp);
    AssignFile(FLigFicSav, Chemin + '\Prospect_d.sav');
    Rewrite (FLigFicSav);

    ProgressionRecup2.Progress := 0;
    ProgressionRecup2.Visible := true;
    screen.cursor:=crHourGlass ;
    if (NombreLignes = 0) then
      NombreLignes := ProgressionTraitement;
    ProgressionRecup2.Maxvalue := NombreLignes;

    ListeDoubles := TStringList.Create;
    while ( not EOF(FDoubleProsp)) do
    begin
      readln (FDoubleProsp,StPro);  //Clé double = T_AUXILIAIRE="01012390"
      ReadTokenPipe(StPro,'=') ;
      ReadTokenPipe(StPro,'=') ;
      CodeProsp := Copy(StPro,2,8);
      ListeDoubles.Add(trim(CodeProsp));
    end;
    ListeDoubles.Sort;  

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PRO') ) do
      readln (FLigProsp,StPro);

    while (Copy(StPro,1,3) = 'PRO') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
         WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'BLO') ) do
      readln (FLigProsp,StPro);

    while (Copy(StPro,1,3) = 'BLO') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'TAB') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'TAB') do
    begin
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'INL') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'INL') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'GRI') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'GRI') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'ACT') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'ACT') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'ACC') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'ACC') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;

    CloseFile(FLigProsp);
    AssignFile(FLigProsp, FichierSav.text);
    Reset (FLigProsp);
    readln (FLigProsp,StPro);
    while ( not EOF(FLigProsp) and (Copy(StPro,1,3)<>'PER') ) do
      readln (FLigProsp,StPro);
    while (Copy(StPro,1,3) = 'PER') do
    begin
      CodeProsp := trim(Copy(StPro,4,10));
      if ( not ListeDoubles.Find(CodeProsp, Index) ) then
      WriteLn (FLigFicSav,StPro) ;
      readln (FLigProsp,StPro);
      ProgressionRecup2.Progress := ProgressionRecup2.Progress+1;
    end;
    BImportGRC.enabled := False;
    bSuivant.Enabled := True;
    ProgressionRecup2.Visible := false;
    CloseFile(FLigProsp);
    CloseFile(FLigFicSav);
    CloseFile(FDoubleProsp);
    ListeDoubles.Free;
    FichierSav.text := Chemin + '\Prospect_d.sav';
    screen.cursor:=crDefault ;
end;

procedure TFAsRecupProsp.Correspondance(var StLig : string) ;
Const MaxPays = 25 ;
Pays : Array [1..MaxPays,1..2] of string[3] =(
('A','DEU'),('AND','AND'),('B','BEL'),('C','COG'),('CAM','CMR'),('CI','CIV'),('E','ESP'),
('ECO','GRB'),('F','FRA'),('GB','GBR'),('GRE','GRC'),('I','ITA'),('L','LUX'),('LI','LEB'),
('LIB','LEB'),('M','MAR'),('MAD','FRA'),('O','FRA'),('PB','NLD'),('POL','POL'),('POR','PRT'),
('S','CHE'),('SEN','SGL'),('T','TUN'),('USA','USA'));
Const MaxJur = 63 ;
Civilite: Array [1..MaxJur,1..2] of string[5] =(
('@ARL','SAR'),('00000','   '),('2ARL','SAR'),('83096','   '),('A','   '),('A.','   '),
('AGA','ASS'),('ARL','SAR'),('ASS','ASS'),('ASSOC','ASS'),('B','   '),('CAB','CAB'),
('CAFET','   '),('CE','ASS'),('CGA','ASS'),('COL','ADM'),('COM S','   '),('COOP','ASS'),
('DELF','   '),('DR','DR '),('E','   '),('ENT','SAR'),('ETB','SAR'),('ETS','SAR'),('EURL','URL'),
('GIA','   '),('GIE','GIE'),('GROUP','SA '),('HOTEL','   '),('INS','   '),('L','   '),
('M','M  '),('ME','MME'),('MELLE','MLE'),('MLLE','MLE'),('MME','MME'),('MR','M  '),('NS','   '),
('R','   '),('REG','   '),('REST','SAR'),('S FA','   '),('S.A','SA '),('S.A.','SA '),('SA','SA '),
('SA.','SA '),('SAEM','   '),('SAR','SAR'),('SARL','SAR'),('SC','SC '),('SCA','SCA'),('SCI','SCI'),('SCM','SCM'),
('SCP','SCP'),('SEC','   '),('SEM','SEM'),('SNC','SNC'),('SSOC','ASS'),('STE','SA '),('TPS','   '),('TRS','   '),
('TS','   '),('URL','URL'));

Juridique : Array [1..MaxJur,1..2] of string[5] =(
('@ARL','SAR'),('00000','   '),('2ARL','SAR'),('83096','   '),('A','   '),('A.','   '),
('AGA','ASS'),('ARL','SAR'),('ASS','ASS'),('ASSOC','ASS'),('B','   '),('CAB','CAB'),
('CAFET','   '),('CE','ASS'),('CGA','ASS'),('COL','ADM'),('COM S','   '),('COOP','ASS'),
('DELF','   '),('DR','   '),('E','   '),('ENT','SAR'),('ETB','SAR'),('ETS','SAR'),('EURL','URL'),
('GIA','   '),('GIE','GIE'),('GROUP','SA '),('HOTEL','   '),('INS','   '),('L','   '),
('M','   '),('ME','   '),('MELLE','   '),('MLLE','   '),('MME','   '),('MR','   '),('NS','   '),
('R','   '),('REG','   '),('REST','SAR'),('S FA','   '),('S.A','SA '),('S.A.','SA '),('SA','SA '),
('SA.','SA '),('SAEM','   '),('SAR','SAR'),('SARL','SAR'),('SC','SC '),('SCA','SCA'),('SCI','SCI'),('SCM','SCM'),
('SCP','SCP'),('SEC','   '),('SEM','SEM'),('SNC','SNC'),('SSOC','ASS'),('STE','SA '),('TPS','   '),('TRS','   '),
('TS','   '),('URL','URL'));



Var StCode1,StCode2,chemin,CodePays: String ;
    i,i_ind1 : Integer ;
    FLigProsp,FLigProsp2: textfile;
    ListeJuri: TStringList;
begin

    // Forme Juridique et Civilité
    StCode1 := Copy(StLig,35,5);
    StCode2 := '     ';
    if (Copy(StLig,298,14) = '00000000000000') then
    begin
      For i:=1 to MaxJur do if (trim(StCode1)=Civilite[i,1]) then StCode2 := Civilite[i,2]+'  ';
    end else
    begin
      For i:=1 to MaxJur do if (trim(StCode1)=Juridique[i,1]) then StCode2 := Juridique[i,2]+'  ';
    end;
    Delete(StLig,35,5) ;
    Insert(StCode2,StLig,35) ;

{  AssignFile(FLigProsp, FichierSav.text);
  Reset (FLigProsp);
  Chemin := UpperCase (ExtractFileDir (FichierSav.text)) ;
  AssignFile(FLigProsp2, Chemin + '\Juridique.sav');
  Rewrite (FLigProsp2);
  while ( not EOF(FLigProsp) and (Copy(StLig,1,3)<>'PRO') ) do
    readln (FLigProsp,StLig);
  ListeJuri := TStringList.Create;
  ListeJuri.Sorted := True;

  while ( not EOF(FLigProsp)) and (Copy(StLig,1,3) = 'PRO') do
  begin
      CodePays := Copy(StLig,35,5);
      ListeJuri.Add(trim(CodePays));
      readln (FLigProsp,StLig);
  end;
  ListeJuri.sort;
  for i_ind1 := 1 to ListeJuri.Count-1 do
    WriteLn (FLigProsp2,ListeJuri.Strings[i_ind1]);
  ListeJuri.free;
      WriteLn (FLigProsp2,StLig) ;
    readln (FLigProsp,StLig);
  end;
  CloseFile(FLigProsp);
  CloseFile(FLigProsp2);
   }
end;

procedure TFAsRecupProsp.CiviliteContact(var StLig : string) ;
Const MaxCiv = 13 ;
Civilite: Array [1..MaxCiv,1..2] of string[5] =(
('DR','DR '),('M','M  '),('MD','MME'),('MDE','MME'),('ME','M  '),('MEL','MLE'),('MELLE','MLE'),
('MLLE','MLE'),('MM','M  '),('MME','MME'),('MME M','MME'),('MP','M  '),('MR','M  '));

Var StCode1,StCode2,chemin,CodePays: String ;
    i,i_ind1 : Integer ;
    FLigProsp,FLigProsp2: textfile;
    ListeJuri: TStringList;
begin

    // Civilité
    StCode1 := Copy(StLig,17,5);
    StCode2 := '     ';
    For i:=1 to MaxCiv do if (trim(StCode1)=Civilite[i,1]) then StCode2 := Civilite[i,2]+'  ';
    Delete(StLig,17,5) ;
    Insert(StCode2,StLig,17) ;
end;

{ types d'actions
  à ne pas remonter dans la GRC :}
function TFAsRecupProsp.ExclureAction(const sttype: string): boolean;
const CodesExclus:string ='012;013;020;021;C05;C06;C07;C09;C14;C16;C20;C21;C22;C23;C24;C32;C33;C35;C39;C42;I01;I02;I03;M07;M08;M09;T06;T24;T25;T26;T31;T33;T35;T38;T41;T42;T43;T44;T58;T59;T60;T61;V16;V17';
begin
   result := false;
   if ( (pos (sttype, CodesExclus)  > 0 )
   or ( (CompareText (sttype,'D01') >= 0) and  (CompareText (sttype,'F99' ) <= 0 ))
   or ( (CompareText (sttype,'M20') >= 0)  and (CompareText (sttype,'M34' ) <= 0 ))
   or ( (CompareText (sttype,'T08') >= 0)  and (CompareText (sttype,'T20' ) <= 0 )))
   then result := true;
end;

{
5 types de perspectives seront maintenues dans la GRC :
001 - PLATEFORME PROSPECT
005 - PLATEFORME CLIENT
007 - PLATEFORME PRESCRIPTION
202 - VD ACTIVITE
20X - SIGNATURE PARTENARIAT  }

function TFAsRecupProsp.ExclurePersp(const sttype: string): integer;
const  CodesNonExclus : string = '001;005;007;202;20X';
begin
   result := pos (sttype, CodesNonExclus);
end;



end.
