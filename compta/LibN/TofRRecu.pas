unit TofRRecu;

interface

uses
    Windows,
    StdCtrls,
    Controls,
    Classes,
    forms,
    ComCtrls,
    Sysutils, // StrToInt, IntToStr, Date, EncodeDate, StrToFloat, TSearchRec, FindFirst, FindNext, faAnyFile, RenameFile, ExtractFilePath, DateToStr, Uppercase
{$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    FE_Main,
    EdtREtat, {LanceEtatTob}
{$ELSE}
    MaineAGL,
    UtileAgl, {LanceEtatTob}
{$ENDIF}
    dpTOFVisuReleve, // CPLanceFicheVisuRel
    Dialogs,  // TOpenDialog
    HCtrls,   // Contrôle Halley
    HEnt1,    // TraduireMemoire, iDate1900, iDate2099, V_PGI
    HMsgBox,  // THMsgBox
    UTOF,     // TOF
    UTOB,     // TOB
    Vierge,   // TFVierge
    Filectrl, // DirectoryExists
    HTB97,    // TToolbarButton97
    Ent1,     // VH
    ParamSoc, // GetParamSocSecur
    EtbUser,  // z_EclatementReleve
    UtilDiv   // CommentExoDateDT
    ;

type TInfoReleve = record
       General,
       Banque,
       Guichet,
       NumCompte,
       Devise,
       DeviseISO,
       RefPointage   : String ;
       DateDu,Dateau : TDateTime ;
       Factor,
       NbrDebit,
       NbrCredit,
       NumReleve     : longint ;
       CumDebit,
       CumCredit     : Double ;
     End ;

procedure CPLanceFiche_RecupReleve;

Type
     {TOF RECUPRELEVE}
     TOF_RecupReleve = Class (TOF)
     private
       BTag       : TButton ;
       BDeTag     : TButton ;
       TagOK : Boolean ;
       TobReleve : TOB;
       GS : THGrid ;
       CBanque : THValComboBox ;
       Repertoire : THedit ;
       CodeDevise : Integer ;
       LMsg : THMsgBox ;
       BHelp: TToolbarButton97;
       Procedure InitGrid ;
       procedure OnChangeCBanque(Sender: TObject) ;
       procedure OnClickBAjoutRel(Sender: TObject) ;
       procedure OnClickBTagOrDeTag(Sender: TObject) ;
       procedure OnClickBChercher(Sender: TObject) ;
       procedure OnClickBReception(Sender: TObject) ;
       procedure OnClickBSupprimer(Sender: TObject) ;
       procedure OnDblClickG(Sender: TObject) ;
       procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
       procedure RempliGrid(RepReleve,CompteBanque  : string);
       procedure RempliSoldeInitial_ESP(TR : TOB ; ReleveCourant : TInfoReleve ; StReleve :  String ) ;
       procedure RempliLigne_ESP(TRLigne : TOB ; StReleve : String ; Var ReleveCourant : TInfoReleve );
       procedure RempliLibelleComplementaire_ESP(TRLigne : TOB ; StReleve : string) ;
       procedure RempliSoldeInitial(TR:TOB;StReleve,General,RefPointage : String;NumReleve:integer) ;
       procedure RempliLigne(TRLigne:TOB;StReleve,General,RefPointage : String;NumReleve,NumLigne:integer);
       procedure RempliSoldeFinal(TR:TOB;StReleve:string;NumLigne:integer);
       procedure RempliLibelleComplementaire(TR,TRLigne:TOB;StReleve:string) ;
       procedure RempliMontant(LaTob : TOB; Montant : Double; NomCredit,NomDebit : string) ;
       procedure RempliMontantEE(LaTob : TOB; Montant : Double; NomCredit,NomDebit : string) ;
       function VerifiReleve(LaDateDeb: TDateTime;General,Devise:string; NomBase : string = '') : Boolean;
       function PutDevise(Devise: string; Fongible: boolean) : integer ;
       function DateReleveOk(DateDeb,DateFin: TDateTime): integer ;
       procedure InitMsg ;
       function AfficheMsg(num : integer;Av,Ap : string ) : Word ;
       procedure BHelpClick(Sender: TObject);
       procedure OnExitRepertoire(Sender: TObject);
     public
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
       function OnIntegreReleve_ESP(Nomreleve : string;Efface:Boolean): boolean;
       function OnIntegreReleve(Nomreleve : string;Efface:Boolean): boolean;
       procedure ImprimerClick(Sender : TObject);
       procedure RemplitTobImp(Fic : string; var T : TOB);
     END ;

const   {Grid Releve}
        SR_NOMFICHIER  = 0 ;
        SR_RIB         = SR_NOMFICHIER + 1 ;
        SR_DU          = SR_RIB + 1 ;
        SR_AU          = SR_DU + 1 ;
        SR_DEVISEINI   = SR_AU + 1 ;


implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    , Constantes, Commun
    , Math, UtilPGI
    ;

procedure CPLanceFiche_RecupReleve;
begin
  if CEstPointageEnConsultationSurDossier then
  begin
    PgiInfo('Vous avez indiqué une liaison avec une comptabilité ' +
            RechDom('CPLIENCOMPTABILITE',GetParamSocSecur('SO_CPLIENGAMME',''), False) +
            ' et la gestion du pointage ' + #10 +
            'est effectuée ' + RechDom('CPPOINTAGESX', GetParamSocSecur('SO_CPPOINTAGESX',''), False) + '. ' +
            'Vous n''avez pas accès à cette commande.', 'Intégration des relevés');
    Exit;
  end;
  AGLLanceFiche('CP', 'RLVINTEGRE', '', '', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /
Description .. : Convertit les montants ETABAC3 en double
Mots clefs ... : ETEBAC3;MONTANT;CONVERTION;
*****************************************************************}
function ConvertMontant(MontantChar : string;Deci:integer) : Double ;
var MontantFloat : Double;Lettre : char;divi,i : integer ;
begin
Lettre:=MontantChar[14] ;
divi:=1;  MontantFloat :=  1;
for i:=1 to deci do divi := divi * 10;
case Lettre of '{' :       begin Lettre := '0';   end;
               '}' :       begin MontantFloat := -1; Lettre := '0';   end;
               'A'..'I' :  begin Dec(Lettre, 16); end;
               'J'..'R' :  begin MontantFloat := -1; Dec(Lettre, 25); end;
               end;
Result:=MontantFloat*StrToFloat(copy(MontantChar,1,13)+Lettre)/divi;
end;

{////////////////////////////////////////////}
{                 RecupReleve                }
{////////////////////////////////////////////}
procedure TOF_Recupreleve.OnLoad;
var BAjoutRel,BChercher,BReception,BVoirReleve,BSupprimer : TButton ;
begin
inherited ;
if Ecran<>nil then TFORM(Ecran).OnKeyDown:=OnKeyDown ;
TFVierge(Ecran).HelpContext:=7773000;
TagOk:=False ;
initMsg ;
BAjoutRel:=TButton(GetControl('BAJOUTREL')) ;
BTag:=TButton(GetControl('BTAG')) ;
BDeTag:=TButton(GetControl('BDETAG')) ;
BChercher:=TButton(GetControl('BCHERCHER')) ;
BReception:=TButton(GetControl('BRECEPTION')) ;
BVoirReleve:=TButton(GetControl('BVOIRRELEVE')) ;
BSupprimer:=TButton(GetControl('BDELETE')) ; BSupprimer.Visible:=True ;
//BAjoutRel:=TButton(GetControl('BAJOUTREL')) ;
//BAjoutRel:=TButton(GetControl('BAJOUTREL')) ;
GS:=THGrid(GetControl('GFILES'));
CBanque:=THValCombobox(GetControl('CBANQUE')) ;
Repertoire:=THEdit(GetControl('REPERTOIRE')) ;
BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
if (BAjoutRel<>nil)  and (not Assigned(BAjoutRel.OnClick))  then BAjoutRel.OnClick:=OnClickBAjoutRel;
if (BTag<>nil)       and (not Assigned(BTag.OnClick))       then BTag.OnClick:=OnClickBTagOrDeTag;
if (BDeTag<>nil)     and (not Assigned(BDeTag.OnClick))     then BDeTag.OnClick:=OnClickBTagOrDeTag;
if (BChercher<>nil)  and (not Assigned(BChercher.OnClick))  then BChercher.OnClick:=OnClickBChercher;
if (BReception<>nil) and (not Assigned(BReception.OnClick)) then BReception.OnClick:=OnClickBReception;
if (GS<>nil)         and (not Assigned(GS.OnDblClick))      then GS.OnDblClick:=OnDblClickG;
if (BVoirReleve<>nil)and (not Assigned(BVoirReleve.OnClick))then BVoirReleve.OnClick:=OnDblClickG;
if (BSupprimer<>nil) (*and (not Assigned(BSupprimer.OnClick))*) then BSupprimer.OnClick:=OnClickBSupprimer;
if (CBanque<>nil)    and (not Assigned(CBanque.OnChange))   then CBanque.OnChange:=OnChangeCBanque;
if (Repertoire<>nil) and (not Assigned(Repertoire.OnExit))   then Repertoire.OnExit:=OnExitRepertoire;
if (BHelp <> nil)    and (not Assigned(BHelp.OnClick))      then BHelp.OnClick:=BHelpClick ;
RemplirValCombo('TTBANQUECP',FiltreBanqueCp('', tcb_bancaire, ''),'',CBanque.Items,CBanque.Values,False,False) ;
if (CBanque<>nil) then begin
   CBanque.Items.Insert(0,'<<Tous>>') ;
   CBanque.Values.Insert(0,'Tous') ;
   CBanque.ItemIndex:=0;
end;
//GV le 15/02/02
if (Not Vh^.OldTeletrans) or (VH^.PaysLocalisation=CodeISOES) then //XVI 24/02/2005
  begin
  SetControlenabled('REPERTOIRE', True) ;
  SetControlEnabled('REPERTLABEL',True) ;
  SetControlText('REPERTOIRE',GetParamSocSecur('SO_REPETEBAC3','')) ;
  end else
  begin
  SetControlenabled('REPERTOIRE', False) ;
  SetControlEnabled('REPERTLABEL',False) ;
  end ;
//Fin gv le 15/02/02
if (BTag<>nil) then BTag.Visible:=True ;
if (BDeTag<>nil) and (BTag<>nil) then BDeTag.Visible:=Not BTag.Visible ;
If (BReception<>NIL) And (Not VH^.OldTeleTrans) Then BReception.Visible:=FALSE ;
OnClickBChercher(nil) ;
  {JP 21/09/04 : FQ TRESO 10120 : possibilité d'imprimer plusieurs relevés}
  SetControlVisible('BIMPRIMER', True); // 14757    EstComptaTreso);
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := ImprimerClick;
end;

procedure TOF_Recupreleve.OnUpdate ;
var i : integer ; Okok: boolean;
Begin
if GS.nbSelected=0 then begin AfficheMsg(9,'','') ; Exit ; end ;
if AfficheMsg(6,'','')<>mrYes then Exit ;
Okok:=False ;
for i := 1 to GS.RowCount-1 do begin
  if (GS.Cells[SR_NOMFICHIER,i]<>'') and GS.IsSelected(i) then begin
    if VH^.PaysLocalisation=CodeISOES then
       Okok:=OnIntegreReleve_ESP(GS.Cells[SR_NOMFICHIER,i],True)
    else //XVI 24/02/2005
       Okok:=OnIntegreReleve(GS.Cells[SR_NOMFICHIER,i],True) ;
    GS.FlipSelection(i) ;
    end ;
  end ;
if Okok then AfficheMsg(7,'','') ;
OnClickBChercher(nil) ;
end;

procedure TOF_Recupreleve.OnClickBSupprimer ;
var i: integer ;Okok: Boolean ;
begin
if GS.nbSelected=0 then begin AfficheMsg(9,'','') ; Exit ; end ;
if AfficheMsg(8,'','')<>mrYes then Exit ;
Okok:=False ;
for i := 1 to GS.RowCount-1 do begin
  if (GS.Cells[SR_NOMFICHIER,i]<>'') and GS.IsSelected(i) then begin
    Okok:=DeleteFile(GS.Cells[SR_NOMFICHIER,i]) ;
    GS.FlipSelection(i) ;
    end ;
  end ;
OnClickBChercher(nil) ;
if Okok then AfficheMsg(16,'','') ;
end;

procedure TOF_Recupreleve.OnClose ;
begin
  LMsg.Free ;
end;

procedure TOF_Recupreleve.InitGrid ;
var i : integer ;
begin
GS.RowCount:=2;
GS.FixedRows:=1 ;
GS.FColAligns[SR_RIB]:=taCenter ;
GS.FColAligns[SR_DU]:=taCenter ;
GS.FColAligns[SR_AU]:=taCenter ;
GS.FColAligns[SR_DEVISEINI]:=taCenter ;
for i:=0 to GS.ColCount-1 do GS.Cells[i,1]:= '' ;
end;

procedure TOF_Recupreleve.RempliGrid(RepReleve,CompteBanque : string);
var MySearchRec: TSearchRec;
    FSource: Textfile ;
    StReleve,NomReleve,Devise : string;
  QQ : TQuery ;
  Okok: Boolean;
  SQL : string ;
  Banque,Guichet,NumCompte,DateDu : String ;
  DateAu : String ;
  ch     : string; {JP 26/07/05 : FQ 16162}
Begin
InitGrid ;
Okok:=(FindFirst(RepReleve, faAnyFile, MySearchRec)=0) ;
try
  if Okok then
  repeat
    NomReleve:=ExtractFilePath(RepReleve)+MySearchRec.Name ;
    AssignFile(FSource,NomReleve);
    Reset(FSource);
    try
      if Not eof(FSource) then begin
        Readln(FSource,StReleve) ;
        Okok:=True;
        if VH^.PaysLocalisation=CodeISOES then Begin
          Banque:=copy(StReleve,3,4) ;
          Guichet:=copy(StReleve,7,4) ;
          Devise:=copy(StReleve,48,3) ;
          NumCompte:=copy(StReleve,11,10) ;
          DateDu:=datetostr(ConvertDate(Copy(streleve,21,6),'AAMMDD')) ;
          DateAu:=datetostr(ConvertDate(Copy(streleve,27,6),'AAMMDD')) ;
        End
        else Begin
          Banque:=copy(StReleve,3,5) ;
          Guichet:=copy(StReleve,12,5) ;
          Devise:=copy(StReleve,17,3) ;
          NumCompte:=copy(StReleve,22,11) ;
          {JP 26/07/05 : FQ 16162 : pour éviter si un relevé n'est pas correct que tous les suivants
                         ne soient pas lus à cause d'un plantage}
          ch := copy(StReleve,35,6);
          if IsNumeric(ch) then DateDu := DateToStr(ConvertDate(ch))
                           else Continue;
        End ;
        if (CompteBanque<>'Tous') then begin
          sql:='SELECT BQ_GENERAL, D_CODEISO FROM BANQUECP LEFT JOIN DEVISE ON D_DEVISE=BQ_DEVISE '
            +' WHERE BQ_ETABBQ="'+copy(StReleve,3,5)+'" AND BQ_GUICHET="'+copy(StReleve,12,5)+'"'
            +' AND BQ_NUMEROCOMPTE="'+copy(StReleve,22,11)+'" AND BQ_GENERAL="'+CompteBanque+'"' ;
          QQ:=OpenSql(SQL,True);

          try
            Okok:=False ;
            while not QQ.Eof do begin
               if QQ.Fields[1].AsString=Devise then Okok:=True ;
               if (not VH^.TenueEuro) and (Devise='EUR') and (QQ.Fields[1].AsString='FRF') then Okok:=True ;
               QQ.Next ;
               if Okok then break ;
              end ;
          finally
            ferme(QQ);
            end ;
          end;
        if Okok then begin
          GS.Cells[SR_NOMFICHIER,GS.RowCount-1]:=NomReleve;
          GS.Cells[SR_RIB,GS.RowCount-1]       :=Banque {copy(StReleve,3,5)}+' '+Guichet{copy(StReleve,12,5)}+' '+NumCompte{copy(StReleve,22,11)} ;
          GS.Cells[SR_DU,GS.RowCount-1]        :=DateDu ; {DateToStr(ConvertDate(copy(StReleve,35,6))) ;}{Date}
          GS.Cells[SR_DEVISEINI,GS.RowCount-1] :=Devise ; {copy(StReleve,17,3);}
          if VH^.PaysLocalisation<>CodeISOES then Begin
              while (copy(StReleve,1,2)<>'07') and (Not eof(FSource)) do Readln(FSource,StReleve) ;
              if copy(StReleve,1,2)='07' then begin
                {JP 26/07/05 : FQ 16162 : pour éviter si un relevé n'est pas correct que tous les suivants
                               ne soient pas lus à cause d'un plantage}
                ch := copy(StReleve,35,6);
                if IsNumeric(ch) then GS.Cells[SR_AU,GS.RowCount - 1] := DateToStr(ConvertDate(ch))
                                 else Continue;
              end
              else
                GS.Cells[SR_AU,GS.RowCount - 1] := '  /  /    ';
          End ;
          GS.RowCount:=GS.RowCount+1 ;
          end;
        end;
    finally
      CloseFile(FSource);
    end ;
  until FindNext(MySearchRec)<>0 ;
  if GS.RowCount<>2 then GS.RowCount:=GS.RowCount-1;
finally
  FindClose(MySearchRec) ;
  end ;
end;

function TOF_Recupreleve.VerifiReleve(LaDateDeb: TDateTime; General,Devise:string; NomBase : string = '') : Boolean;
var QQ:Tquery ;
begin

QQ:=OpenSql('SELECT EE_DATEOLDSOLDE FROM ' + GetTableDossier(NomBase, 'EEXBQ') +' WHERE EE_DATEOLDSOLDE="'+USDateTime(LaDateDeb)+'" AND EE_GENERAL = "'+General+'" AND EE_DEVISE="'+Devise+'"' ,True);
try
  Result:=QQ.Eof ;
finally
  ferme(QQ) ;
  end ;
end;

procedure TOF_RecupReleve.OnExitRepertoire(Sender: TObject) ; // GV le 15/02/02
begin
OnChangeCbanque(Nil) ;
end ;

procedure TOF_RecupReleve.OnChangeCBanque(Sender: TObject) ;
var RepReleve : string ;
Begin
if Not Vh^.OldTeletrans then RepReleve:=getControlText('REPERTOIRE') else RepReleve:=GetParamSocSecur('SO_REPETEBAC3','') ; //GV le 15/02/02
if RepReleve='' then begin AfficheMsg(1,'','') ; exit ; end ;
if RepReleve[length(RepReleve)]<>'\' then RepReleve:=RepReleve+'\' ;
if not DirectoryExists(RepReleve) then
  AfficheMsg(1,'','')
else
  if GS<>nil then RempliGrid(RepReleve+'*.dat',GetControlText('CBANQUE'));
end;

procedure TOF_Recupreleve.OnClickBChercher(Sender: TObject) ;
begin
TagOK:=True ;
OnChangeCBanque(CBanque);
end;

procedure TOF_Recupreleve.OnClickBReception(Sender: TObject) ;
begin
AglLanceFiche('CP','RLVRECEPTION','','','');
end;

procedure TOF_Recupreleve.OnClickBAjoutRel(Sender: TObject) ;
var i : integer ;OpenDialog : TOpenDialog ; Okok: Boolean ;
Begin
Okok:=False ;
//if not VH^.OkModEtebac then begin AfficheMsg(17,'','') ; Exit ; end ; GV OK
if (VH^.PaysLocalisation=CodeISOEs) and ((trim(GetControlText('REPERTOIRE'))='') or (not directoryexists(GetControlText('REPERTOIRE')))) then
 begin
   AfficheMsg(1,'','') ;
   Exit ;
 end ; //XVI 24/02/2005
OpenDialog:=TOpenDialog.create(Ecran);
try
  OpenDialog.InitialDir:=GetControlText('REPERTOIRE') ; // gv le 28/02/02
  if VH^.PaysLocalisation=CodeISOES then begin
     OpenDialog.Options:=OpenDialog.Options+[ofAllowMultiSelect] ;
     OpenDialog.Filter:=TraduireMemoire('Relevés bancaires (*.c43)|*.c43|Tous les fichiers (*.*)|*.*')
  end
  else //XVI 24/02/2005
    OpenDialog.Filter:=TraduireMemoire('Relevés bancaires (*.dat)|*.dat|Tous les fichiers (*.*)|*.*');
  if OpenDialog.execute then begin
    for i:=0 to OpenDialog.Files.Count-1 do begin
      if AfficheMsg(5,OpenDialog.Files.Strings[i],'')=mrYes then
        Okok:=z_EclatementReleve(OpenDialog.Files.Strings[i],GetControlText('REPERTOIRE')) ; // gv 28/02/02
//        Okok:=z_EclatementReleve(OpenDialog.Files.Strings[i],GetParamSocSecur('SO_REPETEBAC3','')) ;
      end;
    end;
finally
  if VH^.PaysLocalisation=CodeISOES then
  Begin
     if OpenDialog.Files.Count<>0 then
        OnChangeCBanque(CBanque) ;
     if not Okok then
       AfficheMsg(14,'','') ;
  End else
  Begin
    if Okok then OnChangeCBanque(CBanque)
    else if OpenDialog.Files.Count <> 0 then
      AfficheMsg(14,'','') ;
  End ; //XVI 24/02/2005
  OpenDialog.free;
end ;
end;

procedure TOF_Recupreleve.OnClickBTagOrDeTag(Sender: TObject) ;
var i: integer;
begin
if not TagOK then Exit ;
for i:=1 to GS.RowCount-1 do begin
  if BDeTag.Visible and GS.IsSelected(i) then GS.FlipSelection(i) ;
  if BTag.Visible   and not GS.IsSelected(i) then GS.FlipSelection(i) ;
  end ;

if uppercase(TComponent(Sender).Name) = 'BTAG' then begin
  BTag.Visible   := false;
  BDeTag.Visible := true;
end;
if uppercase(TComponent(Sender).Name) = 'BDETAG' then begin
  BTag.Visible   := true;
  BDeTag.Visible := false;
end;
end ;

procedure TOF_Recupreleve.OnDblClickG(Sender: TObject) ;
var Argument,Virgule : string ;i: integer ;
Begin
Argument:='NOMFICHIER=' ;
Virgule:='' ;
for i:=1 to GS.RowCount-1 do begin
  Argument:=Argument+Virgule+GS.Cells[SR_NOMFICHIER,i] ;
  Virgule:=',' ;
  end ;
Argument:=Argument+';NUMFICHIER='+IntToStr(GS.Row) ;
CPLanceFicheVisuRel(Argument); // VL 230205 FQ 14757
end;

procedure TOF_RecupReleve.RempliSoldeInitial_ESP(TR : TOB ; ReleveCourant : TInfoReleve ; StReleve :  String ) ;
var
  Montant : double;
begin
   TR.PutValue('EE_GENERAL',ReleveCourant.General) ;
   TR.PutValue('EE_REFPOINTAGE',ReleveCourant.RefPointage) ;
   TR.PutValue('EE_DATEOLDSOLDE',ReleveCourant.DateDu) ;
   TR.PutValue('EE_NUMRELEVE',ReleveCourant.NumReleve) ;
   TR.PutValue('EE_NUMERO',ReleveCourant.NumReleve);

   Montant:=Valeur(copy(StReleve,34,14)) ;
   if stricomp(pChar(Copy(StReleve,33,1)),'2')=0 then //le credit de la banque est le débit chez nous
      Montant:=-Montant ;

   if ReleveCourant.Devise='EUR' then CodeDevise:=2
     else if ReleveCourant.Devise='FRF' then CodeDevise:=1
        else CodeDevise:=3 ;

   RempliMontantEE(TR,Montant/ReleveCourant.Factor,'EE_OLDSOLDECRE','EE_OLDSOLDEDEB') ;

   TR.PutValue('EE_RIB',ReleveCourant.Banque+ReleveCourant.Guichet+ReleveCourant.NumCompte) ;
   TR.PutValue('EE_DEVISE',ReleveCourant.DeviseISO); //??
   TR.PutValue('EE_ORIGINERELEVE','INT'); //Intégré
   TR.PutValue('EE_DATEINTEGRE',date());
   TR.PutValue('EE_STATUTRELEVE','NON');
end;

procedure TOF_RecupReleve.RempliLigne_ESP(TRLigne : TOB ; StReleve : String ; Var ReleveCourant : TInfoReleve );
var
   Montant : double ;
begin
  TRLigne.PutValue('CEL_GENERAL',ReleveCourant.General) ;
  TRLigne.PutValue('CEL_NUMRELEVE',ReleveCourant.NumReleve);
  TRLigne.PutValue('CEL_NUMLIGNE',TRLigne.GetIndex) ;
  //TRLigne.PutValue('CEL_LIBELLE','');  //à remplir par les infos complementaire....
  TRLigne.PutValue('CEL_RIB',Copy(StReleve,7,4)) ; //?? Guichet d'origine...
  TRLigne.PutValue('CEL_CODEAFB',copy(StReleve,23,2) ) ;//Code opération (commun à toutes les banques ESP)
  TRLigne.PutValue('CEL_NATUREINTERNE',copy(StReleve,25,3) ); //(sub)Code opération  (interna à la banque)
  TRLigne.PutValue('CEL_DATEOPERATION',ConvertDate(copy(StReleve,11,6),'AAMMDD'));
  TRLigne.PutValue('CEL_DATEVALEUR',ConvertDate(copy(StReleve,17,6),'AAMMDDD'));
  TRLigne.PutValue('CEL_REFORIGINE',copy(StReleve,43,10)); //Numéro opération Banque
  TRLigne.PutValue('CEL_REFPIECE',copy(StReleve,53,12)); //Reference 1
  TRLigne.PutValue('CEL_REFPOINTAGE',ReleveCourant.RefPointage);
  TRLigne.PutValue('CEL_IMO','');//??
  TRLigne.PutValue('CEL_EXONERE',0); //??

   Montant:=Valeur(copy(StReleve,29,14)) ;
   if stricomp(pChar(Copy(StReleve,28,1)),'2')=0 then begin//le credit de la banque est le débit chez nous
      ReleveCourant.CumCredit:=ReleveCourant.CumCredit+Montant ;
      inc(ReleveCourant.NbrCredit) ;
      Montant:=-Montant ;
   end
   else begin
      ReleveCourant.CumDebit:=ReleveCourant.CumDebit+Montant ;
      inc(ReleveCourant.NbrDebit) ;
   End ;
  RempliMontant(TRLigne,Montant/ReleveCourant.Factor,'CEL_CREDIT','CEL_DEBIT') ;

end;

procedure TOF_RecupReleve.RempliLibelleComplementaire_ESP(TRLigne : TOB ; StReleve : string) ;
Var
  C1,c2 : String ;
  nbr   : integer ;
begin
  Nbr:=Valeuri(Copy(StReleve,3,2)) ;
  if Nbr<=2 then begin
     C1:='' ; C2:='1' ;
     if Nbr>1 then begin
        C1:='2' ;
        C2:='3' ;
     End ;
     TRLigne.PutValue('CEL_LIBELLE'+C1,Copy(StReleve,5,38)) ;
     TRLigne.PutValue('CEL_LIBELLE'+C2,Copy(StReleve,43,38)) ;
  End ;
end;

procedure TOF_RecupReleve.RempliSoldeInitial(TR:TOB;StReleve,General,RefPointage : String;NumReleve:integer) ;
var Montant : double;
    Devise : String ;
begin
TR.PutValue('EE_GENERAL',General);
TR.PutValue('EE_REFPOINTAGE',RefPointage);
TR.PutValue('EE_DATEOLDSOLDE',ConvertDate(copy(StReleve,35,6)) );
TR.PutValue('EE_NUMRELEVE',NumReleve );
TR.PutValue('EE_NUMERO',NumReleve );
Montant:=ConvertMontant(copy(StReleve,91,14),StrToInt(copy(StReleve,20,1)) );
//gv le 15/02/02
Devise := copy(StReleve,17,3) ;
if Devise='EUR' then CodeDevise:=2
  else if Devise='FRF' then CodeDevise:=1
     else CodeDevise:=3 ;
//Fin gv le 15/02/02
RempliMontantEE(TR,Montant,'EE_OLDSOLDECRE','EE_OLDSOLDEDEB') ;
TR.PutValue('EE_RIB',Copy(StReleve,3,5)+Copy(StReleve,12,5)+Copy(StReleve,22,11));
TR.PutValue('EE_DEVISE',copy(StReleve,17,3));
TR.PutValue('EE_ORIGINERELEVE','INT'); //Intégré
TR.PutValue('EE_DATEINTEGRE',date());
TR.PutValue('EE_STATUTRELEVE','NON');
end;

procedure TOF_RecupReleve.RempliLigne(TRLigne:TOB;StReleve,General,RefPointage : String;NumReleve,NumLigne:integer) ;
var Montant : double; Devise : string ;
begin
TRLigne.PutValue('CEL_GENERAL',General);
TRLigne.PutValue('CEL_NUMRELEVE',NumReleve);
TRLigne.PutValue('CEL_NUMLIGNE',NumLigne );
TRLigne.PutValue('CEL_LIBELLE',copy(StReleve,49,31) );
TRLigne.PutValue('CEL_RIB',Copy(StReleve,3,5)+Copy(StReleve,12,5)+Copy(StReleve,22,11)) ;
TRLigne.PutValue('CEL_CODEAFB',copy(StReleve,33,2) );
TRLigne.PutValue('CEL_DATEOPERATION',ConvertDate(copy(StReleve,35,6)));
TRLigne.PutValue('CEL_DATEVALEUR',ConvertDate(copy(StReleve,43,6)));
TRLigne.PutValue('CEL_REFPIECE',copy(StReleve,82,7));
TRLigne.PutValue('CEL_REFPOINTAGE',RefPointage);
TRLigne.PutValue('CEL_IMO',copy(StReleve,21,1));
if copy(StReleve,89,1)<>' ' then TRLigne.PutValue('CEL_EXONERE',copy(StReleve,89,1));
TRLigne.PutValue('NUMECRITURE',copy(StReleve,82,7));
TRLigne.PutValue('NUMLIBELLE','1');
Montant:=ConvertMontant(copy(StReleve,91,14),StrToInt(copy(StReleve,20,1)));
Devise:=copy(StReleve,17,3);
RempliMontant(TRLigne,Montant,'CEL_CREDIT','CEL_DEBIT') ;
end;

procedure TOF_RecupReleve.RempliSoldeFinal(TR:TOB;StReleve:string;NumLigne:integer) ;
var Montant : double;Devise : string ;
begin
TR.PutValue('EE_NBMVT',NumLigne);
TR.PutValue('EE_DATESOLDE',ConvertDate(copy(StReleve,35,6)) );
Montant:=ConvertMontant(copy(StReleve,91,14),StrToInt(copy(StReleve,20,1)));
TR.Putvalue('EE_DATEPOINTAGE',ConvertDate(copy(StReleve,35,6))) ; // gv 14/05/02
Devise := copy(StReleve,17,3);
//gv 14/05/02
if Devise='EUR' then CodeDevise:=2
  else if Devise='FRF' then CodeDevise:=1
     else CodeDevise:=3 ;
//Fin gv 14/05/02
RempliMontantEE(TR,Montant,'EE_NEWSOLDECRE','EE_NEWSOLDEDEB') ;
end;

procedure TOF_RecupReleve.RempliMontant(LaTob : TOB; Montant : Double; NomCredit,NomDebit : string) ;
var {MontantFrf,}MontantEur,MontantDev : double ;
begin
//MontantFrf:=0 ;
MontantEur:=0 ; MontantDev:=0 ;
case CodeDevise of
//1 : MontantFrf:=Montant ;
2 : MontantEur:=Montant ;
3 : MontantDev:=Montant ;
end;
if (Montant>0) then begin
//  LaTob.PutValue(NomCredit       ,MontantFrf) ;
  LaTob.PutValue(NomCredit+'DEV' ,MontantDev) ;
  LaTob.PutValue(NomCredit+'EURO',MontantEur) ;
  end
  else begin
//  LaTob.PutValue(NomDebit       ,MontantFrf*-1) ;
  LaTob.PutValue(NomDebit+'DEV' ,MontantDev*-1) ;
  LaTob.PutValue(NomDebit+'EURO',MontantEur*-1) ;
  end;
end;

procedure TOF_RecupReleve.RempliLibelleComplementaire(TR,TRLigne:TOB;StReleve:string) ;
var NumLibelle : integer; TCh : TOB;
begin
TCh := TR.FindFirst(['NUMECRITURE'],[copy(StReleve,82,7)],TRUE);
if  Tch=nil  then Tch:=TRLigne ;
NumLibelle := StrToInt(TCh.GetValue('NUMLIBELLE'));
if ( NumLibelle = 1 ) then Tch.PutValue('CEL_LIBELLE1',copy(StReleve,41,41));
if ( NumLibelle = 2 ) then Tch.PutValue('CEL_LIBELLE2',copy(StReleve,41,41));
if ( NumLibelle = 3 ) then Tch.PutValue('CEL_LIBELLE3',copy(StReleve,41,41));
if (NumLibelle = 1) or (NumLibelle = 2) or (NumLibelle = 3) then NumLibelle := NumLibelle + 1;
Tch.PutValue('NUMLIBELLE',IntToStr(NumLibelle));
end;

function TOF_Recupreleve.OnIntegreReleve_ESP(NomReleve : string ; Efface : Boolean): boolean;
var
   FSource                       : Textfile ;
   Sql,StReleve,
   CodeOpe,NewNomReleve, NomBase : string ;
   TR,TRLigne                    : Tob ;
   QQ                            : TQuery ;
   Okok                          : boolean ;
   ReleveCourant                 : TInfoReleve ;
begin
   TobReleve:=NIL ;
   try
     AssignFile(FSource,NomReleve); Reset(FSource);
     TobReleve:=TOB.create('_RELEVE',nil,-1) ;
     TR := nil ; TRLigne := nil ;
     okok:=TRUE ;
     //lecture du fichier
     While (Not EOF(FSource)) and (okok) do begin
       Readln(FSource,StReleve) ;
       CodeOpe := copy(StReleve,1,2) ;
       //Gestion de la ligne de fin de Compte :
       if (CodeOpe='33') then begin
          if (not Assigned(TR)) then begin
             AfficheMsg(18,'','') ;
             okok:=FALSE ;
          End else
          if (Valeuri(Copy(StReleve,21,5))<>ReleveCourant.NbrDebit) or (Valeur(Copy(StReleve,26,14))<>ReleveCourant.CumDebit) then begin
             AfficheMsg(19,'','') ;
             okok:=FALSE ;
          end else
          if (Valeuri(Copy(StReleve,40,5))<>ReleveCourant.NbrCredit) or (Valeur(Copy(StReleve,45,14))<>ReleveCourant.CumCredit) then begin
             AfficheMsg(20,'','') ;
             okok:=FALSE ;
          end
          else begin
             {$IFDEF TRESO}
             if not InsertOrUpdateAllNivelsMS(TobReleve, NomBase, True) then begin
                AfficheMsg(13,'','') ;
                Okok:=False ;
             end ;
             {$ELSE}
             if not TobReleve.InsertDB(nil,True) then begin
                AfficheMsg(13,'','') ;
                Okok:=False ;
             end ;
             {$ENDIF TRESO}
             TObReleve.ClearDetail ;
             TRLigne:=nil ;
             TR:=Nil ;
          End ;
       End Else
       //Gestion de la ligne de Entete de Compte : Ancien Solde
       if (CodeOpe='11') then begin
          if (Assigned(TR)) then begin
             AfficheMsg(18,'','') ;
             okok:=FALSE ;
          end
          else begin
             Fillchar(ReleveCourant,sizeof(TInfoReleve),#0) ;
             ReleveCourant.Banque:=copy(StReleve,3,4)      ; ReleveCourant.Guichet:=copy(StReleve,7,4) ;
             ReleveCourant.NumCompte:=copy(StReleve,11,10) ; ReleveCourant.DeviseISO:=Copy(StReleve,48,3) ;
             ReleveCourant.DateDu:=ConvertDate(Copy(StReleve,21,6),'AAMMDD') ;
             ReleveCourant.DateAu:=ConvertDate(Copy(StReleve,27,6),'AAMMDD') ;
             {$IFDEF TRESO}
             QQ:=OpenSQL('SELECT BQ_GENERAL, DOS_NOMBASE FROM BANQUECP ' +
                        'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                        'WHERE BQ_ETABBQ="'+ReleveCourant.Banque+'" AND BQ_GUICHET="'+ReleveCourant.Guichet+'" AND BQ_NUMEROCOMPTE="'+ReleveCourant.NumCompte+'"',true);
             {$ELSE}
             QQ:=OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_ETABBQ="'+ReleveCourant.Banque+'" AND BQ_GUICHET="'+ReleveCourant.Guichet+'" AND BQ_NUMEROCOMPTE="'+ReleveCourant.NumCompte+'"',true);
             {$ENDIF TRESO}
             if Not QQ.EOF then begin
                {$IFDEF TRESO}
                NomBase := QQ.FindField('DOS_NOMBASE').AsString;
                {$ELSE}
                NomBase := '';
                {$ENDIF TRESO}
                ReleveCourant.General:=QQ.Fields[0].AsString ;
                //Verification des dates de relevés puis intégration dans la base !
                case DateReleveOk(ReleveCourant.DateDu,ReleveCourant.Dateau) of
                  0: begin AfficheMsg(10,'','') ; Okok:=False ; end ;
                  1: begin AfficheMsg(11,'','') ; Okok:=False ; end ;
                  2: begin AfficheMsg(12,'','') ; Okok:=False ; end ;
                  3: if Not VerifiReleve(ReleveCourant.DateDu,ReleveCourant.General,ReleveCourant.DeviseISO, NomBase) then begin
                        AfficheMsg(2,'','') ;
                        Okok:=False ;
                     end ;
                end ;
             End
             else Begin
                AfficheMsg(3,'','');
                Okok:=False ;
             End ;
             Ferme(QQ) ;
          End ;
          if okok then begin
             Sql:='SELECT max(EE_NUMERO) FROM ' + GetTableDossier(NomBase, 'EEXBQ') + ' WHERE EE_GENERAL = "'+ReleveCourant.General+'"' ;
             QQ:=OpenSQL(Sql,true) ;
             ReleveCourant.numReleve:=QQ.Fields[0].AsInteger ;
             Ferme(QQ) ;
             Sql:='SELECT max(CEL_NUMRELEVE) FROM ' + GetTableDossier(NomBase, 'EEXBQLIG') + ' WHERE CEL_GENERAL = "'+ReleveCourant.General+'"' ;
             QQ:=OpenSQL(Sql,true) ;
             ReleveCourant.numReleve:=maxintvalue([ReleveCourant.numReleve,QQ.Fields[0].AsInteger]) ;
             Ferme(QQ) ;
             QQ:=OpenSQL('SELECT D_DEVISE, D_DECIMALE FROM DEVISE WHERE D_CODEISO="'+ReleveCourant.DeviseISO+'"' ,True);
             CodeDevise:=-1 ;
             if QQ.eof then begin
                AfficheMsg(15,'','');
                Okok:=False ;
             end
             else Begin
                ReleveCourant.Devise:=QQ.Fields[0].AsString ;
                ReleveCourant.Factor:=round(Power(10,QQ.Fields[1].AsInteger)) ;
                {CodeDevise:=PutDevise(QQ.Fields[1].AsString,(QQ.Fields[3].AsString='X')) ;
                if (not VH^.TenueEuro) and (ReleveCourant.Devise='EUR') and (QQ.Fields[2].AsString='FRF') and (CodeDevise<0) then
                   CodeDevise:=2 ;
                if (CodeDevise<0) then begin
                   AfficheMsg(15,'','');
                   Okok:=False ;
                end ;}
             End ;
             ferme(QQ);
          End ;

         if Okok then begin
           inc(ReleveCourant.NumReleve) ;
           ReleveCourant.RefPointage:=Uppercase(copy(BourreLaDoncSurLesComptes(ReleveCourant.General),1,12)+'R'+IntToStr(ReleveCourant.NumReleve)) ;
           TR:=TOB.Create('EEXBQ',TobReleve,-1);
           RempliSoldeInitial_ESP(TR,ReleveCourant,StReleve) ;
         end;
       end else
       //Gestion des lignes d'écritures
       if (CodeOpe = '22') then begin
          if (not Assigned(TR)) then begin
             AfficheMsg(18,'','') ;
             okok:=FALSE ;
          end
          else begin
             TRLigne:=TOB.Create('EEXBQLIG',TR,-1);
             RempliLigne_ESP(TRLigne,StReleve,Relevecourant);
          end;
       End else
       //Gestion des lignes de commentaires supplementaires
       if (CodeOpe = '23') and (Assigned(TRLigne)) then
          RempliLibelleComplementaire_ESP(TRLigne,StReleve) else
       //Gestion de la ligne de montant DEVISE
       if (CodeOpe = '24') and (Assigned(TRLigne))  then begin
          //Rien à faire (on n'a pas des champs pour stocker l'information....)
       end ;
     end ;
    finally
    CloseFile(FSource);
    if Assigned(TOBReleve) then
       TobReleve.free;
    end ;
   //Suppression des relevés intégrés !
   if okok then Begin
      if Not Efface then if AfficheMsg(4,NomReleve,'')=mrYes then Efface:=TRUE ;
      if Efface then begin
         NewNomReleve:=ChangeFileExt(NomReleve,'.INT') ;
         RenameFile(NomReleve,NewNomReleve) ;
         end ;
      End ;
   result:=Okok ;
end;

function TOF_Recupreleve.OnIntegreReleve(NomReleve : string;Efface:Boolean): boolean;
var FSource: Textfile;
    Sql,StReleve,Rib,CodeOpe,Banque,Guichet,Compte,General,Devise,RefPointage,NewNomReleve,NomFic : string ;
    TR,TRLigne: Tob;
    Q: TQuery;
    NumReleve,NumLigne: integer;
    DateDeb,DateFin: TDateTime ;
    Okok: boolean ;
  NumReleveSolde, NumreleveLigne : Integer ;
  PremiereFois : Boolean ;
  NomBase : string;
begin
DateDeb:=IDate1900; DateFin:=IDate2099; Okok:=True ; PremiereFois:=True ;
TobReleve := TOB.create('_RELEVE',nil,-1) ;

try
  AssignFile(FSource,NomReleve); Reset(FSource);
  try
    TR := nil ; TRLigne := nil ; NumLigne := 0 ; NumReleve:=0 ;
(*
lecture du fichier
*)
    While Not EOF(FSource) do
    begin
      Readln(FSource,StReleve) ;
      CodeOpe := copy(StReleve,1,2) ;
(*
Gestion de la ligne de solde Entete : Ancien Solde
*)
      if ( CodeOpe = '01' )  then
      begin
        Banque:=copy(StReleve,3,5) ;      Guichet:=copy(StReleve,12,5) ;
        Compte:=copy(StReleve,22,11) ;    Rib:=Banque+Guichet+Compte;
        //gv
        {$IFDEF TRESO}
        Q:=OpenSQL('SELECT BQ_GENERAL, DOS_NOMBASE FROM BANQUECP ' +
                   'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                   'WHERE BQ_ETABBQ="'+Banque+'" AND BQ_GUICHET="'+Guichet+'" AND BQ_NUMEROCOMPTE="'+Compte+'"',true);
        {$ELSE}
        Q:=OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_ETABBQ="'+Banque+'" AND BQ_GUICHET="'+Guichet+'" AND BQ_NUMEROCOMPTE="'+Compte+'"',true);
        {$ENDIF TRESO}
        if Not Q.EOF then
          begin
          {$IFDEF TRESO}
          NomBase := Q.FindField('DOS_NOMBASE').AsString;
          {$ELSE}
          NomBase := '';
          {$ENDIF TRESO}
          General:=Q.Fields[0].AsString ;
          Ferme(Q) ;
          // gv le 31/05/02 recherche du + grand numero de releve entre eexbq et eexbqlig .. pb de cle en dble
          if premiereFois then  // gv 03/06/02
            begin
            NumReleveSolde:=0 ; NumReleveLigne:=0 ;
            Sql:='SELECT max(EE_NUMERO) FROM ' + GetTableDossier(NomBase, 'EEXBQ') + ' WHERE EE_GENERAL = "'+General+'"' ;
            Q:=OpenSQL(Sql,true) ;
            if not Q.EOF then NumReleveSolde:=Q.Fields[0].AsInteger ;
            Ferme(Q) ;

            Sql:='SELECT max(CEL_NUMRELEVE) FROM ' + GetTableDossier(NomBase, 'EEXBQLIG') + ' WHERE CEL_GENERAL = "'+General+'"' ;
            Q:=OpenSQL(Sql,true) ;
            If Not Q.EOF then NumReleveLigne:=Q.Fields[0].AsInteger ;
            Ferme(Q) ;
            if (NumReleveLigne>0) or (NumReleveSolde>0) then
              if NumReleveSolde >= NumReleveLigne then NumReleve:=NumReleveSolde else NumReleve:=NumReleveLigne
              else NumReleve:=0 ;
            PremiereFois:=False ;
            end ;
            // fin gv le 31/05/02
          end else
          begin
          AfficheMsg(3,'','');
              Result:=False ;
          ferme(Q) ;
          exit ;
          end ;
        Q:=OpenSQL('SELECT D_LIBELLE, D_DEVISE, D_CODEISO,D_FONGIBLE FROM DEVISE WHERE D_CODEISO="'+Copy(StReleve,17,3)+'"' ,True);
        try
          CodeDevise:=-1 ;
          if Q.eof then begin AfficheMsg(15,'',''); Okok:=False ; end ;
          Devise:=copy(StReleve,17,3) ;
          while not Q.Eof do
            begin
            if Q.Fields[2].AsString=Devise then CodeDevise:=PutDevise(Q.Fields[1].AsString,(Q.Fields[3].AsString='X'));
            if (not VH^.TenueEuro) and (Devise='EUR') and (Q.Fields[2].AsString='FRF') and (CodeDevise<0) then CodeDevise:=2 ;
            Q.Next ;
            end ; // fin du while not Q.EOF
           if (CodeDevise<0) and Okok then begin AfficheMsg(15,'',''); Okok:=False ; end ;
        finally
          ferme(Q);
        end ;
        // fin gv

        if Okok<>False then
        begin
          // gv le 05/03/02
          inc(NumReleve) ;
          // fin gv le 05/03/02
          //RefPointage := Uppercase(General+TraduireMemoire('REL')+IntToStr(NumReleve)) ;

          //RR 28032003 règle le pb d'un cpte général >= 12 et un champ=17 car.
          //donc on limite le cpt à 12, d'où 12 + R + 0000 = 17 car.
          RefPointage := Uppercase(copy(BourreLaDoncSurLesComptes(General),1,12)+'R'+IntToStr(NumReleve)) ;
          TR := TOB.Create('EEXBQ',TobReleve,-1);
          RempliSoldeInitial(TR,StReleve,General,RefPointage,NumReleve) ;
          DateDeb:=TR.GetValue('EE_DATEOLDSOLDE') ;
        end;
      end ;
(*
Gestion des lignes d'écritures
*)
      if (CodeOpe = '04') and (TR <> nil)  then
        begin
        TRLigne := TOB.Create('EEXBQLIG',TR,-1);
        TRLigne.AddChampSup('NUMECRITURE',False);
        TRLigne.AddChampSup('NUMLIBELLE',False);
        Inc(NumLigne);
        RempliLigne(TRLigne,StReleve,General,RefPointage,NumReleve,Numligne);
        end;
(*
Gestion des lignes de commentaires supplementaires
*)
      if (CodeOpe = '05') and (TRLigne <> nil)  then RempliLibelleComplementaire(TR,TRLigne,StReleve);
(*
Gestion de la ligne de Nouveau Solde
*)
      if (CodeOpe = '07') and (TR <>nil)  then
        begin
        RempliSoldeFinal(TR,StReleve,NumLigne) ;
        DateFin:=TR.GetValue('EE_DATESOLDE') ;
        end ;
    end ;
  finally
  CloseFile(FSource);
  end ;
(*
Verification des dates de relevés puis intégration dans la base !
*)

  if Okok then
  begin
    case DateReleveOk(DateDeb,DateFin) of
      0: begin AfficheMsg(10,'','') ; Okok:=False ; end ;
      1: begin AfficheMsg(11,'','') ; Okok:=False ; end ;
      2: begin AfficheMsg(12,'','') ; Okok:=False ; end ;
      {$IFDEF TRESO}
      3: if not VerifiReleve(DateDeb,General,Devise, NomBase) then begin
           AfficheMsg(2, '', '');
           Okok:=False;
         end
         else if not InsertOrUpdateAllNivelsMS(TobReleve, NomBase, True) then begin
           AfficheMsg(13,'','');
           Okok:=False;
         end;
      {$ELSE}
      3: if Not VerifiReleve(DateDeb,General,Devise) then begin AfficheMsg(2,'','') ; Okok:=False ; end else
           if not TobReleve.InsertDB(nil,True) then begin AfficheMsg(13,'','') ; Okok:=False ; end;
      {$ENDIF TRESO}
    end ;
  end ;
finally
TobReleve.free;
end ;
(*
Suppression des relevés intégrés !
*)
Efface:=Okok ;
if Not Efface then if AfficheMsg(4,NomReleve,'')=mrYes then Efface:=TRUE ;
//if Efface then DeleteFile(NomReleve) ;
if Efface then // GV  Rename des fichiers integrés
  begin
  NomFic:=NomReleve ;
  NewNomReleve:=ReadTokenPipe(NomFic,'.') ;
  NewNomReleve:=NewNomReleve+'.INT' ;
  RenameFile(NomReleve,NewNomReleve) ;
  end ; //fin gv
result:=Okok ;
end;

function TOF_Recupreleve.DateReleveOk(DateDeb,DateFin: TDateTime): integer ;
begin
result:=CommentExoDateDT(DateDeb) ;
if result=3 then result:=CommentExoDateDT(DateFin) ;
end ;

function TOF_Recupreleve.PutDevise(Devise: string; Fongible: boolean) : integer ;
begin
if not VH^.TenueEuro then begin
  if Devise='EUR' then Result:=2
    else if V_PGI.DevisePivot=Devise then Result:=1
           else Result:=3
  end
  else begin
  if Devise='EUR' then Result:=1
    else if Fongible then Result:=2
           else Result:=3
  end ;
end;

// PFUGIER: 22/12//2000
// Un grand moment ce code
procedure TOF_Recupreleve.InitMsg ;
begin
LMsg:=THMsgBox.create(FMenuG) ;
{00}LMsg.Mess.Add('1;Intégration des relevés;Erreur Inconnue;W;O;O;O');
{01}LMsg.Mess.Add('1;Intégration des relevés;Le nom du répertoire ETEBAC3 n''est pas valide;W;O;O;O'); // GV le 15/02/02
{02}LMsg.Mess.Add('1;Intégration des relevés;Le relevé existe déjà, vous ne pouvez pas l''integrer;W;O;O;O');
{03}LMsg.Mess.Add('1;Intégration des relevés;Aucune banque pour ce relevé;W;O;O;O');
{04}LMsg.Mess.Add('1;Intégration des relevés;Voulez-vous effacer le relevé %% ?;W;YN;N;N');
{05}LMsg.Mess.Add('1;Intégration des relevés;Voulez-vous ajouter le relevé %% ?;Q;YN;N;N');
{06}LMsg.Mess.Add('1;Intégration des relevés;Voulez vous intégrer les relevés sélectionnés ?;Q;YN;N;N');
{07}LMsg.Mess.Add('1;Intégration des relevés;L''intégration des relevés est terminée;I;O;O;O');
{08}LMsg.Mess.Add('1;Intégration des relevés;Voulez-vous effacer tous les fichiers sélectionnés;W;YN;N;N');
{09}LMsg.Mess.Add('1;Intégration des relevés;Aucun fichier sélectionné;W;O;O;O');
{10}LMsg.Mess.Add('1;Intégration des relevés;Les dates du relevé ne correspondent pas à un exercice de la base;W;O;O;O');
{11}LMsg.Mess.Add('1;Intégration des relevés;Les dates du relevé correspondent à un exercice cloturé;W;O;O;O');
{12}LMsg.Mess.Add('1;Intégration des relevés;Les dates du relevé correspondent à une période cloturée;W;O;O;O');
{13}LMsg.Mess.Add('1;Intégration des relevés;Erreur lors de l''intégration du relevé dans la base;W;O;O;O');
if VH^.PaysLocalisation=CodeISOEs then
   {14}LMsg.Mess.Add('1;Intégration des relevés;Attention ! Il y a des fichiers dont le format n''est pas bon, ces fichiers n''ont pas été intégrés;W;O;O;O')
else
    {14}LMsg.Mess.Add('1;Intégration des relevés;Attention ! Vous ne pouvez pas intégrer ce fichier ce n''est pas un relevé bancaire;W;O;O;O'); //XVI 24/02/2005
{15}LMsg.Mess.Add('1;Intégration des relevés;Devise incorrecte. Veuillez vérifier le paramétrage des codes devises et des codes Iso;W;O;O;O');
{16}LMsg.Mess.Add('1;Intégration des relevés;La suppression des relevés est terminée;I;O;O;O');
{17}LMsg.Mess.Add('1;Intégration des relevés;Le module Liaison Bancaire n''est pas sérialisé ;W;O;O;O;');
{18}LMsg.Mess.Add('1;Intégration des relevés;Le format du fichier n''est pas bon!;W;O;O;O;');
{19}LMsg.Mess.Add('1;Intégration des relevés;Erreur d''intégration, différences sur le débit!;W;O;O;O;');
{20}LMsg.Mess.Add('1;Intégration des relevés;Erreur d''intégration, différences sur le crédit!;W;O;O;O;');
end;

function TOF_Recupreleve.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
begin
Result:=mrNone ;  if LMsg=nil then exit ;
if (Num<0) or (Num>LMsg.Mess.Count-1) then Num:=0 ; //erreur inconnue
Result:=LMsg.Execute(num,Av,Ap) ;
end;

procedure TOF_Recupreleve.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if key=VK_F9 then OnClickBChercher(nil) ;
end ;

procedure TOF_Recupreleve.BHelpClick(Sender: TObject);
begin CallHelpTopic(Ecran) ; end;

{JP 21/09/04 : FQ TRESO 10120 : possibilité d'imprimer plusieurs relevés
{---------------------------------------------------------------------------------------}
procedure TOF_Recupreleve.ImprimerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  {Si aucune ligne sélectionnées}
  if (GS.nbSelected = 0) and not GS.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un relevé.;W;O;O;O;', '', '');
    Exit;
  end;

  T := TOB.Create('$$$', nil, -1);
  try
    {Constitution de la tob d'impression}
    for n := 0 to GS.RowCount - 1 do
      if GS.IsSelected(n)then
        RemplitTobImp(GS.Cells[SR_NOMFICHIER, n], T);
    {Lancement de l'état}
    LanceEtatTob('E', 'ECT', 'REL', T, True, False, False, nil, '', 'Intégration de relevés bancaires', False);
  finally
    FreeAndNil(T);
  end;
end;

{JP 21/09/04 : FQ TRESO 10120 : Remplit la Tob d'impression à partir d'un fichier de relevé
{---------------------------------------------------------------------------------------}
procedure TOF_Recupreleve.RemplitTobImp(Fic : string; var T : TOB);
{---------------------------------------------------------------------------------------}

    {----------------------------------------------------------------------------}
    function ConvertMontant(MontantChar : string; Deci : Integer) : Double ;
    {----------------------------------------------------------------------------}
    var
      MontantFloat : Double;
      Lettre : Char;
      Divi   : Integer;
      n      : Integer;
    begin
      Lettre := MontantChar[14] ;
      Divi := 1;
      MontantFloat :=  1;

      for n := 1 to Deci do Divi := Divi * 10;

      case Lettre of
        '{'      : Lettre := '0';
        '}'      : begin
                     MontantFloat := -1;
                     Lettre := '0';
                   end;
        'A'..'I' : Dec(Lettre, 16);
        'J'..'R' : begin
                     MontantFloat := -1;
                     Dec(Lettre, 25);
                   end;
      end;
      Result := MontantFloat * StrToFloat(Copy(MontantChar, 1, 13) + Lettre) / Divi;
    end;

    {----------------------------------------------------------------------------}
    function ConvertDate(DateChar : string) : TDateTime ;
    {----------------------------------------------------------------------------}
    var
      a, m, j : Word;
    begin
      a  := StrToInt(Copy(DateChar, 5, 2));
      if a > 90 then a := a + 1900
                else a := a + 2000;
      m := StrToInt(Copy(DateChar, 3, 2));
      j := StrToInt(Copy(DateChar, 1, 2));
      Result:= EncodeDate(a, m, j);
    end;

var
  D : TOB;
  F : TextFile;
  Q : TQuery;
  Ligne : string;
  Code  : string;
  Cpte  : string;
  Deci  : Integer;
  Montant : Double;
begin
  if (Fic = '') or not Assigned(T) then Exit;

  AssignFile(F, Fic);
  Reset(F);
  Cpte := '';

  try
    while not EOF(F) do begin
      Readln(F, Ligne);
      {Recherche du libellé du compte à partir du RIB}
      if Cpte = '' then begin
        Q := OpenSQL('SELECT BQ_LIBELLE FROM BANQUECP WHERE BQ_ETABBQ = "' + Copy(Ligne, 3, 5) + '" AND ' +
                     'BQ_GUICHET = "' + Copy(Ligne, 12, 5) + '" AND BQ_NUMEROCOMPTE = "' + Copy(Ligne, 22, 11) + '"', True);
        if not Q.EOF then Cpte := Q.FindField('BQ_LIBELLE').AsString
                     else Cpte := '    ';
        Ferme(Q);
      end;

      Code := Copy(Ligne, 1, 2);
      if (Code = '01') or (Code = '04') or (Code = '07') then begin
        D := TOB.Create('$$$', T, -1);
        {Gestion du type de ligne, du libellé du compte et de son RIB}
        D.AddChampSupValeur('CODE', Code);
        D.AddChampSupValeur('COMPTE', Cpte);
        D.AddChampSupValeur('RIB', Copy(Ligne, 3, 5) + ' ' + Copy(Ligne, 12, 5) + ' ' + Copy(Ligne, 22, 11));

        {Gestion des champs DEBIT et CREDIT}
        Deci := StrToInt(Copy(Ligne, 20, 1));
        Montant := ConvertMontant(Copy(Ligne, 91, 14), Deci);
        D.AddChampSup('DEBIT', False);
        D.AddChampSup('CREDIT', False);
             if Montant < 0 then D.SetString('DEBIT' , FloatToStrF(Montant * -1, ffFixed, 20, Deci))
        else if Montant > 0 then D.SetString('CREDIT', FloatToStrF(Montant, ffFixed, 20, Deci))
        else if Montant = 0 then D.SetString('CREDIT', '0.00');

        {Ajout de la date du relevé et de la devise}
        D.AddChampSupValeur('DATERELEV', DateToStr(ConvertDate(Copy(Ligne, 35, 6))));
        D.AddChampSupValeur('DEVISE', Copy(Ligne, 17, 3));

        {Gestion du libellé des lignes et de la date de valeur}
        D.AddChampSup('DATEVALEUR', False);
             if (Code = '01') then D.AddChampSupValeur('LIBELLE', 'Solde Initial ')
        else if (Code = '07') then D.AddChampSupValeur('LIBELLE', 'Solde Final ')
        else begin
          D.AddChampSupValeur('LIBELLE', Copy(Ligne, 49, 31));
          D.SetString('DATEVALEUR', DateToStr(ConvertDate(Copy(Ligne, 43, 6))));
        end;
        end;
      end ;
  finally
    CloseFile(F);
  end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 31/03/2005
Modifié le ... :   /  /    
Description .. :
Suite ........ : FQ 15593 : Suite passage à l'Euro, les champs qui ont été
Suite ........ : conservés ne sont plas les même entre les tables EEXBQ et
Suite ........ : EESBQLIG, d'ou besoin de traiter diffrement les champs
Suite ........ : EE_NEWSOLDExxx / EE_OLDSOLDExxx et les champs
Suite ........ : CEL_DEBITxxx / CEL_CREDITxxx
Mots clefs ... :
*****************************************************************}
procedure TOF_RecupReleve.RempliMontantEE(LaTob: TOB; Montant: Double; NomCredit, NomDebit: string);
var MontantEur,MontantDev : double ;
begin
MontantEur:=0 ; MontantDev:=0 ;
case CodeDevise of
2 : MontantEur:=Montant ;
3 : MontantDev:=Montant ;
end;
if (Montant>0) then begin
  LaTob.PutValue(NomCredit ,MontantDev) ;
  LaTob.PutValue(NomCredit+'EURO',MontantEur) ;
  end
  else begin
  LaTob.PutValue(NomDebit ,MontantDev*-1) ;
  LaTob.PutValue(NomDebit+'EURO',MontantEur*-1) ;
  end;
end;

Initialization
    registerclasses([TOF_RECUPRELEVE]);

end.
