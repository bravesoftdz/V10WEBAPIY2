{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 08/02/2001
Modifié le ... : 10/07/2001
Description .. : Edition différée des documents
Mots clefs ... : EDITION;DIFFEREE;DOCUMENT
*****************************************************************}
unit UTofEditDocDiff;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,uTofAfBaseCodeAffaire,
{$IFDEF EAGLCLIENT}
      UtileAGL,eMul,MaineAGL,
{$ELSE}
      dbTables,Fiche, mul, db,DBGrids,   Fe_Main,EdtEtat, EdtDoc,
{$IFNDEF V530}
      EdtREtat, EdtRDoc,
{$ENDIF}
{$ENDIF}
{$IFDEF BTP}
      BTPUtil,
{$ENDIF}
{$IFDEF NOMADE}
      UtilPOP,
{$ENDIF}
{$IFDEF MISESOUSPLI}
      uPDFBatch,
{$ENDIF}
      Printers,
{$IFDEF AFFAIRE}
    DicoAf,AffaireUtil,
{$ENDIF}
      {$IFNDEF V530} uRecupSQLModele, {$ENDIF}
      M3VM, M3FP, Hqry, EntGC, FactComm, FactUtil, Ent1, UtilMulTrt, Paramsoc,UtilGC,
      FactTob ;

Type
     TOF_EditDocDiff_Mul = Class (TOF_AFBASECODEAFFAIRE)
     Private
        titre : string;
{$IFDEF MISESOUSPLI}
        PrintFileName : string ;
        PrintToFile : Boolean ;
        PrinterIndex : Integer ;
        FFSortie : TextFile ;
        Function  AffichePrintDialog : boolean ;
        Function  OuvreFichierSortie ( NomPDF : String ) : Boolean ;
        Function  AppendFichierSortie ( NomPDFTemp : String ) : Boolean ;
        Function  CloseFichierSortie  : Boolean ;
{$ENDIF}
        procedure AlimPieceVivante;
    public
        procedure OnArgument (stArgument : string); override;
        function ChercheCriteresDeTri : string;
        procedure EditeLesDocs (TobPiece : TOB);
        procedure LanceEdition ;
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
        Function  NomCorrect ( NomPDF : String ) : Boolean ;
        procedure RenseignePiece (TobP : TOB);
    end ;

procedure AGLEditDocDiff_mul(parms:array of variant; nb: integer ) ;

const
// libellés des messages
TexteMessage: array[1..3] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'Changement d''imprimé.' + #13 + 'Les documents suivants s''éditeront d''après le modèle '
          {2} ,'Trop de documents sélectionnés. Veuiller réduire la sélection'
              );

implementation

uses Dialogs;

procedure TOF_EditDocDiff_Mul.OnArgument (stArgument : string) ;
var F : TFMul;
    stPlus : string;
    CC     : THValComboBox ;
    stLesArg : string; // DBR rassemblement
begin
// DBR Rassemblement
{$IFDEF CCS3}
if ctxGCAFF in V_PGI.PGIContexte then
  begin
  TTabSheet(GetControl('PRESCLT')).TabVisible := False;
  SetControlVisible('GP_AFFAIRE',False);
  SetControlVisible('GP_AFFAIRE1',False); SetControlVisible('GP_AFFAIRE2',False);
  SetControlVisible('GP_AFFAIRE3',False); SetControlVisible('GP_AVENANT',False);
  SetControlVisible('BSELECTAFF1',False); SetControlVisible('TGP_AFFAIRE1',False);
  SetControlVisible('BEFFACEAFF1',False);
  end;
{$ENDIF}
{$IFDEF AFFAIRE}
    if (ctxAffaire in V_PGI.PGIContexte) then
    begin
      F:=TFMul(Ecran);
      F.DBListe := 'AFMULPIECEEDIT';
      if F.Q <> NIL then F.Q.Liste  := 'AFMULPIECEEDIT';
      SetControlProperty ('AFF_GROUPECONF', 'DataType', 'YYGROUPECONF');  //MCD 28/10/03 pas mis dans la fiche !....
      THValComboBox(GetControl ('TRIEDITION1')).Plus  := 'and CO_CODE like "G%" ORDER BY CO_TYPE, CO_CODE';
      THValComboBox(GetControl ('TRIEDITION2')).Plus  := 'and CO_CODE like "G%" ORDER BY CO_TYPE, CO_CODE';
      THValComboBox(GetControl ('TRIEDITION3')).Plus  := 'and CO_CODE like "G%" ORDER BY CO_TYPE, CO_CODE';

      TTabSheet(GetControl('PAFFAIRE')).TabVisible := True;
      TTabSheet(GetControl('PSTATAFF')).TabVisible := True;
    end else
{$ENDIF}
    begin
      SetControlProperty ('TRIEDITION1', 'DataType', 'GCRUPTURE');
      SetControlProperty ('TRIEDITION2', 'DataType', 'GCRUPTURE');
      SetControlProperty ('TRIEDITION3', 'DataType', 'GCRUPTURE');
      THValComboBox(GetControl ('TRIEDITION1')).Plus  := ' ORDER BY CO_TYPE, CO_CODE';
      THValComboBox(GetControl ('TRIEDITION2')).Plus  := ' ORDER BY CO_TYPE, CO_CODE';
      THValComboBox(GetControl ('TRIEDITION3')).Plus  := ' ORDER BY CO_TYPE, CO_CODE';
      TTabSheet(GetControl('PAFFAIRE')).TabVisible := False;
      TTabSheet(GetControl('PSTATAFF')).TabVisible := False;
    end;
    stLesArg := stArgument;
    stArgument := ReadTokenSt (stLesArg);
    if StArgument = 'STOCK' then
      StPlus:='AND (GPP_NATUREPIECEG="TEM" OR GPP_NATUREPIECEG="TRE" OR GPP_NATUREPIECEG="EEX" OR GPP_NATUREPIECEG="SEX")'
    else if StArgument <> '' then
      stplus := 'AND GPP_VENTEACHAT="' + stArgument + '"';
    {$IFDEF NOMADE}
      if stArgument = 'VEN' then
        stPlus := stPlus + GetNaturePOP('GPP_NATUREPIECEG');
    {$ENDIF} // NOMADE
{$IFDEF MISESOUSPLI}
    TTabSheet(GetControl('PAFFAIRE')).TabVisible := False;
    TTabSheet(GetControl('PSTATAFF')).TabVisible := False;
    THValComboBox(GetControl ('TRIEDITION1')).Plus  := 'and CO_CODE not like "X%" and CO_LIBRE like "GP_%"';
    THValComboBox(GetControl ('TRIEDITION2')).Plus  := 'and CO_CODE not like "X%" and CO_LIBRE like "GP_%"';
    THValComboBox(GetControl ('TRIEDITION3')).Plus  := 'and CO_CODE not like "X%" and CO_LIBRE like "GP_%"';
{$ELSE}
{$IFNDEF AFFAIRE}
(*    THValComboBox(GetControl ('TRIEDITION1')).Plus  := 'and CO_CODE not like "X%" and CO_LIBRE like "GP_%"';
    THValComboBox(GetControl ('TRIEDITION2')).Plus  := 'and CO_CODE not like "X%" and CO_LIBRE like "GP_%"';
    THValComboBox(GetControl ('TRIEDITION3')).Plus  := 'and CO_CODE not like "X%" and CO_LIBRE like "GP_%"'; *)

    THValComboBox(GetControl ('TRIEDITION1')).Plus  := 'ORDER BY CO_TYPE, CO_CODE';
    THValComboBox(GetControl ('TRIEDITION2')).Plus  := 'ORDER BY CO_TYPE, CO_CODE';
    THValComboBox(GetControl ('TRIEDITION3')).Plus  := 'ORDER BY CO_TYPE, CO_CODE';

    THValComboBox(GetControl ('TRIEDITION1')).DataType := 'GCRUPTURE';
    THValComboBox(GetControl ('TRIEDITION2')).DataType := 'GCRUPTURE';
    THValComboBox(GetControl ('TRIEDITION3')).DataType := 'GCRUPTURE';
{$ENDIF}
{$ENDIF}

inherited ;

//*** Blocages des natures de pièces autorisées en affaires
if (ctxAffaire in V_PGI.PGIContexte) then
   begin
     if (stArgument='ACH') then
     begin
       stplus:=stPlus + AfPlusNatureAchat;
 //mcd à supprimer      SetControlText('GP_NATUREPIECEG','FF');
     end
     else if (CtxScot in V_PGI.PGIContexte) then stplus := Stplus +' AND (GPP_NATUREPIECEG="FPR" OR GPP_NATUREPIECEG="APR" OR GPP_NATUREPIECEG="FAC" OR GPP_NATUREPIECEG="AVC" OR GPP_NATUREPIECEG="FRE")';
     SetControlVisible ('GP_DEPOT',false);
     SetControlVisible ('TGP_DEPOT',false);
     SetControlVisible ('GP_DOMAINE',false);
     SetControlVisible ('TGP_DOMAINE',false);
   end;
// **** Fin affaire ***

{$IFDEF AFFAIRE} // DBR rassemblement
if ctxAffaire in V_PGI.PGIContexte then
begin
  titre := Ecran.caption;
  SetControlChecked('GP_EDITEE', false);
end;
SetControlVisible ('BIMPFORMULE',GetParamSoc('SO_AFVARIABLES'));
SetControlVisible ('BIMPREVISION',GetParamSoc('SO_AFREVISIONPRIX'));
{$ENDIF} // AFFAIRE

//mcd if not(ctxScot in V_PGI.PGIContexte) then setControlproperty ('GP_NATUREPIECEG','Plus',stplus);
setControlproperty ('GP_NATUREPIECEG','Plus',stplus);

if stArgument = 'VEN' then
    begin
    if ctxAffaire in V_PGI.PGIContexte then
      SetControlText ('GP_NATUREPIECEG', 'FAC')
    else
       //mcd ajoute test scot
      if ctxscot in V_Pgi.PgiContexte then SetControlText('GP_NATUREPIECEG','FF')
      else SetControlText ('GP_NATUREPIECEG', 'CC');
    SetControlCaption ('TGP_TIERS', 'Client de');
    end else
    begin
    if Not (ctxScot in V_PGI.PGIcontexte) then SetControlText ('GP_NATUREPIECEG', 'CF')
      else SetControlText ('GP_NATUREPIECEG', 'FF');
    setControlproperty ('GP_TIERS','DataType','GCTIERSFOURN');
    setControlproperty ('GP_TIERS_','DataType','GCTIERSFOURN');
    SetControlText ('TGP_TIERS', 'Fournisseur de');
    end;
CC:=THValComboBox(GetControl('GP_DOMAINE'))       ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
{$IFDEF BTP}
SetControlVisible('PCOMPLEMENT',False);
{$ENDIF}
if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
   begin
   SetControlVisible ('TGP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE1', False);
   SetControlVisible ('GP_AFFAIRE2', False);
   SetControlVisible ('GP_AFFAIRE3', False);
   SetControlVisible ('GP_AVENANT', False);
   end;
//if not VH_GC.GCMultiDepots then SetControlCaption('TGP_DEPOTDEST','Etablis. destinataire');
if not VH_GC.GCMultiDepots then SetControlCaption('TGP_DEPOT','Etablissement');
end;

Procedure TOF_EditDocDiff_Mul.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ; Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ; Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
END ;

procedure TOF_EditDocDiff_Mul.LanceEdition ;
var F : TFMul ;
    TobPiece : TOB ;
//    iInd : integer ;
//    stWhere, stNature : string ;
//    TSql : TQuery ;
    stSQL, stTableCompteur : string ; // DBR Rassemblement
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
    HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
    exit;
    end;

{$IFDEF AFFAIRE} // DBR rassemblement
if ctxAffaire in V_PGI.PGIContexte then
begin
    If (PGIAskAf ('Confirmez-vous l''édition de ces factures', titre) <> mrYes) then exit;
end;
{$ENDIF} // AFFAIRE DBR

TobPiece := TOB.Create ('', Nil, -1) ;
if TobPiece = nil then exit ;

// pour utilisation TraiteEnregMulTable
{$IFDEF MISESOUSPLI}
stSql:='SELECT PIECE.*,GPA_CODEPOSTAL FROM PIECE LEFT JOIN PIECEADRESSE ADR1 ON ADR1.GPA_NATUREPIECEG=GP_NATUREPIECEG ';
stSql:=stSql+' AND ADR1.GPA_SOUCHE=GP_SOUCHE AND ADR1.GPA_NUMERO=GP_NUMERO AND ADR1.GPA_INDICEG=GP_INDICEG ';
stSql:=stSql+' AND ADR1.GPA_NUMLIGNE=0 and ((GP_NUMADRESSEFACT=2 and ADR1.GPA_TYPEPIECEADR="002") or (GP_NUMADRESSEFACT=1 and ADR1.GPA_TYPEPIECEADR="001")) ';
stTableCompteur := 'PIECE';
{$ELSE} // MISESOUPLI
{$IFDEF AFFAIRE}
//stSql := 'SELECT P.* FROM PIECE P LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE';
//stTableCompteur := 'AFPIECEAFFAIRE'; AB-20031028
stSql := 'SELECT P.* FROM PIECE P LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE LEFT OUTER JOIN TIERS TI ON GP_TIERS=T_TIERS';
stTableCompteur := 'PIECE LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE LEFT OUTER JOIN TIERS ON GP_TIERS=T_TIERS ';
{$ELSE} // AFFAIRE
stSql := 'SELECT * FROM PIECE ';
stTableCompteur := 'PIECE';
{$ENDIF} // AFFAIRE
{$ENDIF} // MISESOUSPLI
TraiteEnregMulTable (TFMul(Ecran), stSql,
                     'GP_NATUREPIECEG;GP_SOUCHE;GP_NUMERO;GP_INDICEG', 'PIECE',
                     'GP_NUMERO', stTableCompteur, TobPiece, True, 2000);

if TobPiece.Detail.Count > 0 then
    begin
    EditeLesDocs (TobPiece) ;
    if GetControlText ('GP_EDITEE') = '-' then
        TobPiece.UpdateDBTable(True);
    end;

if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
F.bSelectAll.Down := False ;
TobPiece.free ;
end;

Function TOF_EditDocDiff_Mul.NomCorrect ( NomPDF : String ) : Boolean ;
Var FFText : TextFile ;
BEGIN
Result:=False ;
if NomPDF='' then Exit ;
  if not VH_GC.GCIfDefCEGID then
if Pos('.PDF',NomPDF)<=0 then Exit ;
AssignFile(FFText,NomPDF) ; {$i-} Rewrite(FFText) ; {$i+}
if IoResult=0 then
   BEGIN
   CloseFile(FFText) ; DeleteFile(NomPDF) ;
   Result:=True ;
   END ;
END ;

{$IFDEF MISESOUSPLI}
Function TOF_EditDocDiff_Mul.OuvreFichierSortie ( NomPDF : String ) : Boolean ;
BEGIN
Result:=False ;
if NomPDF='' then Exit ;
AssignFile(FFSortie,NomPDF) ; {$i-} Rewrite(FFSortie) ; {$i+}
if IoResult=0 then
   BEGIN
   Result:=True ;
   END ;
END ;

Function TOF_EditDocDiff_Mul.AppendFichierSortie ( NomPDFTemp : String ) : Boolean ;
var TempFic : TextFile ;
    st : string;
BEGIN
Result:=true;
If nompdfTemp='' then exit;
AssignFile (TempFic, NomPdfTemp);
Reset (TempFic);
while not EOF(TempFic) do
    begin
    readln (TempFic,st);
    writeln(FFSortie, st);
    end;
closeFile( TempFic );
DeleteFile (NomPdfTemp);
END;

Function TOF_EditDocDiff_Mul.CloseFichierSortie  : Boolean ;
BEGIN
Result:=true;
CloseFile (FFSortie);
END;
{$ENDIF}

function TOF_EditDocDiff_Mul.ChercheCriteresDeTri : string;
var TriEdit1, TriEdit2, TriEdit3 : THValComboBox;
begin
Result := '';
TriEdit1 := THValComboBox(GetControl('TRIEDITION1'));
TriEdit2 := THValComboBox(GetControl('TRIEDITION2'));
TriEdit3 := THValComboBox(GetControl('TRIEDITION3'));
if (TriEdit1<>nil) then
    if (TriEdit1.Value<>'') and (TriEdit1.Value<>TraduireMemoire('<<Aucun>>')) then
    	Result := TriEdit1.Value;

if (TriEdit2<>nil) then
    if (TriEdit2.Value<>'') and (TriEdit2.Value<>TraduireMemoire('<<Aucun>>')) then
    begin
        if Result <> '' then Result:=Result+';'+TriEdit2.Value
        else Result := TriEdit2.Value;
    end;

if (TriEdit3<>nil) then
    if (TriEdit3.Value<>'') and (TriEdit3.Value<>TraduireMemoire('<<Aucun>>')) then
    begin
        if Result <> '' then Result:=Result+';'+TriEdit3.Value
        else Result := TriEdit3.Value;
    end;


if (pos('GP_NUMERO', Result)=0) then
    begin
    if Result <> '' then Result := Result+';GP_NUMERO'
    else Result := 'GP_NUMERO';
    end;
end;

procedure TOF_EditDocDiff_Mul.EditeLesDocs (TobPiece : TOB) ;
var iInd, iNbExemplaire,iNbPiece : integer;
    CleDoc : R_CleDoc;
    stSql,stsqlVar, stCle, stModele,NomPDF,SQLOrder : string;
    stModeleSuivant, stTri : string;
    Bdupli : string; // AFFAIRE
    TL : TList;
    TT : TStrings;
    BApercu, BAp, BDuplicata, BFichierPDF, OldQRThread : boolean;
    CBFichierPDF : TCheckBox ;
    ENomPDF     : THEdit ;
//    Pages : TpageControl;
    iPrint : integer; // depuis la 560c
{$IFDEF MISESOUSPLI}
    NomPDFTemp : string ;
    MemePli : Boolean ;
{$ENDIF}
    ChoixImp : boolean;  // DBR Rassemblement
    st : string;
begin
//Pages:=Nil;
BApercu := TCheckBox(GetControl ('BAPERCU')).Checked;
//BDuplicata := TCheckBox(GetControl ('BDUPLICATA')).Checked; AFFAIRE DEBUT
case TCheckBox(GetControl('BDUPLICATA')).State of
    cbGrayed   : Bdupli := '|' ;
    cbChecked  : Bdupli := 'X' ;
    cbUnChecked: Bdupli := '-' ;
    end;
{$IFDEF AFFAIRE} 
if (ctxAffaire in V_PGI.PGIContexte) then
begin
  if GetParamSoc('SO_AFVARIABLES') then //Affaire-ONYX
  begin
    if TCheckBox(GetControl('BIMPFORMULE')).Checked then
         TobPiece.Detail[0].AddChampSupValeur ('BIMPFORMULE', 'X', True)
    else TobPiece.Detail[0].AddChampSupValeur ('BIMPFORMULE', '-', True);
  end;
  if GetParamSoc('SO_AFREVISIONPRIX') then //Affaire-ONYX
  begin
    if TCheckBox(GetControl('BIMPREVISION')).Checked then
         TobPiece.Detail[0].AddChampSupValeur ('BIMPREVISION', 'X', true)
    else TobPiece.Detail[0].AddChampSupValeur ('BIMPREVISION', '-', true);
  end;
end;
{$ENDIF} // AFFAIRE FIN


TobPiece.Detail[0].AddChampSupValeur ('ETATSTANDARD_IMPMODELE', '', True);
TobPiece.Detail[0].AddChampSupValeur ('NBEXEMPLAIRE', 0, true);
TobPiece.Detail[0].AddChampSup ('OKETAT', True);
for iInd := 0 to TobPiece.Detail.count - 1 do
    begin
{$IFNDEF MISESOUSPLI}
        if ((TCheckBox(GetControl('CBFORCEMODELE')).Checked)) and
           (THValComboBox(GetControl('MODELEEDITION')).value <> '') then // DBR Rassemblement Début
        Begin
            TobPiece.Detail[iInd].PutValue ('NBEXEMPLAIRE', '1');
            TobPiece.Detail[iInd].PutValue ('ETATSTANDARD_IMPMODELE', '0' + THValComboBox(GetControl('MODELEEDITION')).value);
            if (TRadioButton(GetControl('RBEDITETAT')).Checked = true) then
                TobPiece.Detail[iInd].PutValue ('OKETAT', 'X')
            else
                TobPiece.Detail[iInd].PutValue ('OKETAT', '-');  // DBR Rassemblement Fin
        end else
{$ENDIF}
            RenseignePiece (TobPiece.Detail[iInd]);
    end;

{Gestion PDF}
BFichierPDF:=False ;
CBFichierPDF := TCheckBox(GetControl ('BFICHIERPDF')) ;
if CBFichierPDF <> Nil then BFichierPDF := CBFichierPDF.Checked;
OldQRThread := V_PGI.QRMultiThread ;
NomPDF:='' ;
if BFichierPDF then
begin
    ENomPDF:=THEdit(GetControl('NOMPDF'));
    if ENomPDF<>Nil then NomPDF:=uppercase(ENomPDF.Text) ;
end;

//tri par modéle   0:modéle par défaut, 1 :modéle CLient/FAC et 2 modéle du parpiéce
//  pourquoi ce tri sur l'origine du modèle ?  DBR Rassemblement : Pour les changements d'imprimé
stTri := ChercheCriteresDeTri; // DBR Rassemblement

{$IFDEF MISESOUSPLI}
TobPiece.Detail.Sort ('ETATSTANDARD_IMPMODELE;GPA_CODEPOSTAL;GP_TIERSFACTURE;GP_DATEPIECE;GP_NUMERO');
{$ELSE}
TobPiece.Detail.Sort ('ETATSTANDARD_IMPMODELE;' + stTri);
{$ENDIF}

{$IFDEF AFFAIRE}
ChoixImp:=True; // DBR Rassemblement / AFFAIRE : print dialogue sur la premiere edition uniquement
{$ELSE}
ChoixImp:=False;
{$ENDIF}

{$IFDEF MISESOUSPLI}
  if VH_GC.GCIfDefCEGID then
  begin
    if Not OuvreFichierSortie ( NomPDF ) then begin  PGIError ('Nom de fichier incorrect');exit; end;
    PrintToFile := True;
    NomPDFTemp:=ChangeFileExt(NomPdf,'.tmp');
    PrintFileName := NomPDFTemp;
    if not AffichePrintDialog then Exit;
    V_PGI.NoPrintDialog:=True ;
  end
 else
  begin
    if Not NomCorrect(NomPDF) then NomPDF:='' else
    BEGIN
      V_PGI.QRPDF:=True ; V_PGI.QRPDFQueue:=NomPDF ; V_PGI.QRPDFMerge:='' ;
      V_PGI.QRMultiThread:=False ;
    END ;
    v_pgi.MiseSousPli :=TRUE ;
    v_pgi.PageMiseSousPli:=0 ;
    if NomPDF<>'' then StartPDFBatch(NomPdf) ;
  end;
{$ELSE}
  if BFichierPDF then
  begin
    if Not NomCorrect(NomPDF) then NomPDF:='' else
    begin
      V_PGI.QRPDF:=True ; V_PGI.QRPDFQueue:=NomPDF ; V_PGI.QRPDFMerge:='' ;
      V_PGI.QRMultiThread:=False ;
    end;
  end;
{$ENDIF}

iPrint := 0;

for iInd := 0 to TobPiece.Detail.Count - 1 do
    begin
    stModele := copy (TobPiece.Detail[iInd].GetValue ('ETATSTANDARD_IMPMODELE'), 2,
                      length (TobPiece.Detail[iInd].GetValue ('ETATSTANDARD_IMPMODELE')) - 1);

    stCle := TobPiece.Detail[iInd].GetValue ('GP_NATUREPIECEG') + ';' +
                        DateToStr (TobPiece.Detail[iInd].GetValue ('GP_DATEPIECE')) + ';' +
                        TobPiece.Detail[iInd].GetValue ('GP_SOUCHE') + ';' +
                        IntToStr (TobPiece.Detail[iInd].GetValue ('GP_NUMERO')) + ';' +
                        IntToStr (TobPiece.Detail[iInd].GetValue ('GP_INDICEG'));
    StringToCleDoc (stCle, CleDoc);

{$IFDEF CCS3}
    if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then
       begin
       stSql:=RecupSQLModele('E','GPJ', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False));
       if (pos('ORDER BY',uppercase(stSql))<=0) then stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE' ;
//       Pages := TPageControl.Create(Application);
       end else
       begin
       stSql:=RecupSQLModele('L','GPI', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False));
       if (pos('ORDER BY',uppercase(stSql))<=0) then stSQL := stSQL+ ' ORDER BY GL_NUMLIGNE';
       end ;
{$ELSE}
    if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then
       begin
       SQLOrder:='';
       stSql:=RecupSQLEtat('E','GPJ', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False),SQLOrder);
       if not (ctxaffaire in V_PGI.PGICOntexte) then stSql:=stSql+' and GL_NONIMPRIMABLE<>"X"' ;  //mcd 23/10/03 mis dans les requêtes de nos états. car on peut avoir un état qui n'imrpiem que les non imprimable
       if (SQLOrder='') then stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE'
                        else stSql:=stSql+' '+SQLOrder;
//       Pages := TPageControl.Create(Application);
       END else
       BEGIN
       SQLOrder:='';
       stSql:=RecupSQLEtat('L','GPI', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False),SQLOrder);
       if (SQLOrder='') then stSql := stSql+ ' ORDER BY GL_NUMLIGNE'
                        else stSql:=stSql+' '+SQLOrder;
       END;
{$ENDIF}

    if NomPDF='' then BAp:=BApercu
    else
{$IFDEF MISESOUSPLI}
      if VH_GC.GCIfDefCEGID then
        BAp:=False
      else
        BAp:=True ;
{$ELSE}
      BAp:=True ;
{$ENDIF}
    iNbPiece := StrToInt(TobPiece.Detail[iInd].GetValue ('NBEXEMPLAIRE'));
    if iNbPiece = 0 then iNbPiece :=1;
{$IFDEF MISESOUSPLI}
    if not VH_GC.GCIfDefCEGID then
    // si pas rupture sur client on force même enveloppe
    if (iInd<(TobPiece.Detail.Count-1))
           and (TobPiece.Detail[iInd].GetValue ('GP_TIERSFACTURE')=TobPiece.Detail[iInd+1].GetValue ('GP_TIERSFACTURE'))
           then MemePli := true else MemePli := False;
{$ENDIF}
    for iNbExemplaire := 1 to iNbpiece do
        begin
        TL := TList.Create ;
        TT := TStringList.Create ;
        TT.Add (stSql);
        TL.Add (TT);
        if ((iInd > 0) or (iNbExemplaire > 1)) and (Bap = False) then V_PGI.NoPrintDialog := True;
{$IFDEF MISESOUSPLI}
        if VH_GC.GCIfDefCEGID then
        begin
          V_PGI.PrintToFile := PrintToFile;
          V_PGI.PrintFileName := PrintFileName;
          V_PGI.NoPrintDialog:= True;
          V_PGI.QRPrinterIndex := PrinterIndex;
        end
        else
          v_pgi.MiseSousPliContinue := MemePli ; // autoreset à false
{$ENDIF}
        if (Bdupli='X') then BDuplicata := True // DBR Rassemblement
        else if (Bdupli='-') then BDuplicata:=False
        else if InbExemplaire=1 then BDuplicata:=False
        else BDuplicata:=true;

        if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then
          BEGIN
{$IFNDEF MISESOUSPLI}
{$IFDEF AFFAIRE}
           V_PGI.NoPrintDialog := Not(ChoixImp);  //mcd 29/01/02
{$ENDIF}
{$ENDIF}
          iPrint := LanceEtat ('E','GPJ',stModele,Bap, false, false,Nil,trim (stSql),'',BDuplicata);
{$IFDEF AFFAIRE}
          if (ctxAffaire in V_PGI.PGIContexte) then
          begin
            if GetParamSoc('SO_AFVARIABLES') and (iPrint > 0) then
            begin
              if (TobPiece.Detail[iInd].GetValue ('BIMPFORMULE') = 'X') then //Affaire-ONYX
              begin
                stSqlVar:=RecupSQLModele('E','AON', 'JUS','','','',' WHERE ACT_NUMPIECE="'+EncodeRefPiece(TobPiece.Detail[iInd])+'" AND '+WherePiece(CleDoc,ttdPiece,False));
                LanceEtat ('E','AON','JUS',Bap, false, false,Nil,trim (stSqlvar),'',BDuplicata);
              end;
            end;
            if GetParamSoc('SO_AFREVISIONPRIX') and (iPrint > 0) then
            begin
              if (TobPiece.Detail[iInd].GetValue ('BIMPREVISION') = 'X') then    //Affaire-ONYX
              begin
                stSqlVar:=RecupSQLModele('E','AON', 'FRE','','','',' WHERE '+WherePiece(CleDoc,ttdLigne,False)+' AND (GL_FORCODE1<>"" OR GL_FORCODE2<>"") ORDER BY GL_NUMLIGNE ' );
                LanceEtat ('E','AON','FRE',Bap, false, false,Nil,trim (stSqlvar),'',BDuplicata);
              end;
            end;
          end;
{$ENDIF}
           END else
           BEGIN
           iPrint := LanceDocument ('L', 'GPI', stModele, TL, Nil, Bap,ChoixImp, BDuplicata);
           END ;
{$IFNDEF MISESOUSPLI} // DBR Rassemblement
{$IFDEF AFFAIRE}
        choiximp := False;
{$ENDIF}
{$ENDIF}
//        BAp:=false;
        TT.Free;
        TL.Free;
        if iPrint <= 0 then break;
{$IFDEF MISESOUSPLI}
        if VH_GC.GCIfDefCEGID then
           AppendFichierSortie(PrintFileName);
{$ENDIF}
        end;

//    if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then Pages.Free;

    if iPrint <= 0 then break;

    if (GetControlText ('BINFORMECHGTETAT') = 'X') and (NomPDF='') and
       (iInd < TobPiece.Detail.Count - 1) then
    begin
        stModeleSuivant := copy (TobPiece.Detail[iInd + 1].GetValue ('ETATSTANDARD_IMPMODELE'), 2,
                                 length (TobPiece.Detail[iInd].GetValue ('ETATSTANDARD_IMPMODELE')) - 1);
        if stModele <> stModeleSuivant then
        begin
//            HShowMessage('0;'+TFMUL(Ecran).Caption+';'+TexteMessage[2]+' '+RechDom ('GCIMPMODELE', stModeleSuivant, False) +';W;O;O;O;', '', '');
            if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then st := format(TexteMessage[2]+ ' : %s', [RechDom ('GCIMPETAT', stModeleSuivant, False)])
            else st := format(TexteMessage[2]+ ' : %s', [RechDom ('GCIMPMODELE', stModeleSuivant, False)]);
            If (PGIAsk(st,titre)<> mrYes) then break; // DBR Rassemblement
        end;
    end;
{$IFDEF MISESOUSPLI}
    TobPiece.detail[iInd].VirtuelleToReelle('PIECE');
{$ENDIF}
    if (GetControlText ('GP_EDITEE') = '-') and (V_PGI.QRPrinted) then
        TobPiece.Detail[iInd].PutValue ('GP_EDITEE', 'X');

{$IFNDEF MISESOUSPLI}
    V_PGI.QRPDFMerge:=NomPDF ;
{$ENDIF}
    end;
{$IFDEF MISESOUSPLI}
    if VH_GC.GCIfDefCEGID then
    begin
      CloseFichierSortie;
      V_PGI.PrintToFile := False;
      V_PGI.PrintFileName := '';
      V_PGI.NoPrintDialog:= False;
    end
    else
    begin
      CancelPDFBatch ;
      v_pgi.MiseSousPli:=False ;
      v_pgi.PageMiseSousPli:=0 ;
    end;
{$ENDIF}
V_PGI.QRPDF:=False ; V_PGI.QRPDFQueue:='' ; V_PGI.QRPDFMerge:='' ; V_PGI.QRMultiThread:=OldQRThread ;
end;

procedure TOF_EditDocDiff_Mul.RenseignePiece (TobP : TOB);
var QTiers : TQuery;
    stEtatStandard, Suff : string;
    NaturePiece, Modele, Domaine, Etab :string;
    NbCopies : integer;
    OkEtat : string;
begin
NaturePiece:=TobP.GetValue('GP_NATUREPIECEG') ;
OkEtat:='-' ;
Etab:=TobP.GetValue('GP_ETABLISSEMENT') ;
Domaine:=TobP.GetValue('GP_DOMAINE') ;
Suff:='' ;
//if TobP.GetValue('GP_SAISIECONTRE')='X' then Suff:='CON' ;
Modele:='' ;
NbCopies:=GetInfoParPiece(NaturePiece,'GPP_NBEXEMPLAIRE') ;
if ((NbCopies<0) or (NbCopies>99)) then NbCopies:=0 ;

{Recherche sur profils Affaires}
{$IFDEF AFFAIRE}
if (TobP.GetValue ('GP_AFFAIRE') <> '') and GetparamSoc('SO_AFPARPIECEAFF') then
begin
  QTiers := OpenSQL('SELECT API_IMPMODELE,API_IMPETAT,API_NBEXEMPLAIRE,API_IMPREVISION,API_IMPFORMULE'
         + ' FROM AFFAIREPIECE  WHERE API_AFFAIRE="' + TobP.GetValue ('GP_AFFAIRE') +'"'
         + ' AND API_NATUREPIECEG="' + TobP.GetValue ('GP_NATUREPIECEG') + '"',false);
  if not QTiers.Eof then
     BEGIN
     Modele:=QTiers.FindField('API_IMPETAT').AsString ;
     if Modele<>'' then OkEtat:='X' else Modele:=QTiers.FindField('API_IMPMODELE').AsString ;
     if (ctxAffaire in V_PGI.PGIContexte) then
     begin
       if GetParamSoc('SO_AFREVISIONPRIX') then
        if (TCheckBox(GetControl('BIMPREVISION')).State = cbGrayed) then  //Affaire-ONYX
         TobP.PutValue ('BIMPREVISION', QTiers.FindField('API_IMPREVISION').AsString);
       if GetParamSoc('SO_AFVARIABLES') then
        if (TCheckBox(GetControl('BIMPFORMULE')).State = cbGrayed) then   //Affaire-ONYX
         TobP.PutValue ('BIMPFORMULE', QTiers.FindField('API_IMPFORMULE').AsString);
     end;
     NbCopies:=QTiers.FindField('API_NBEXEMPLAIRE').AsInteger ; if ((NbCopies<0) or (NbCopies>99)) then NbCopies:=0 ;
     stEtatStandard := '5';
     END ;
  Ferme(QTiers);
end;
{$ENDIF}

{Recherche sur profils tiers}
if Modele='' then
  BEGIN
  QTiers := OpenSQL('SELECT GTP_IMPMODELE' + Suff + ', GTP_IMPETAT' + Suff
         + ', GTP_NBEXEMPLAIRE FROM TIERSPIECE WHERE GTP_TIERS="' + TobP.GetValue ('GP_TIERS')
         + '" AND GTP_NATUREPIECEG="' + TobP.GetValue ('GP_NATUREPIECEG') + '"',false);
  if not QTiers.Eof then
     BEGIN
     Modele:=QTiers.FindField('GTP_IMPETAT'+Suff).AsString ;
     if Modele<>'' then OkEtat:='X' else Modele:=QTiers.FindField('GTP_IMPMODELE'+Suff).AsString ;
     NbCopies:=QTiers.FindField('GTP_NBEXEMPLAIRE').AsInteger ; if ((NbCopies<0) or (NbCopies>99)) then NbCopies:=0 ;
     stEtatStandard := '4';
     END ;
  Ferme(QTiers);
  END;

{Recherche sur "les "PARPIECE}
if Modele='' then
   BEGIN
   {Etat Nature/Domaine}
   stEtatStandard := '3' ;
   Modele:=GetInfoParPieceDomaine(NaturePiece,Domaine,'GDP_IMPETAT'+Suff) ;
   if Modele<>'' then  OkEtat:='X' else
      BEGIN
      {Document Nature/Domaine}
      Modele:=GetInfoParPieceDomaine(NaturePiece,Domaine,'GDP_IMPMODELE'+Suff) ;
      if Modele='' then
         BEGIN
         stEtatStandard := '2' ;
         {Etat Nature/Etablissement}
         Modele:=GetInfoParPieceCompl(NaturePiece,Etab,'GPC_IMPETAT'+Suff) ;
         if Modele<>'' then OkEtat:='X' else
            BEGIN
            {Document Nature/Etablissement}
            Modele:=GetInfoParPieceCompl(NaturePiece,Etab,'GPC_IMPMODELE'+Suff) ;
            if Modele='' then
               BEGIN
               stEtatStandard := '1' ;
              {Etat Nature}
               Modele:=GetInfoParPiece(NaturePiece,'GPP_IMPETAT'+Suff) ;
              {Modèle Nature}
               if Modele<>'' then OkEtat:='X' else Modele:=GetInfoParPiece(NaturePiece,'GPP_IMPMODELE'+Suff) ;
               END ;
            END ;
         END ;
      END ;
   END ;

// Recherche sur Paramsoc
if Modele='' then
begin
    stEtatStandard := '0' ;
    Modele:=VH_GC.GCImpModeleDefaut ;
end;

// --- Modif BTP
// Recherche spéciale pour le modèle des situations de travaux (BTP)
{$IFDEF BTP}
if (NaturePiece='FBT') and (RenvoieTypeFact (TobP.GetValue('GP_AFFAIREDEVIS')) = 'AVA') then
  begin
  // Récupérer le modèle des situations dans les paramètres
  Modele:=GetParamsoc('SO_BTETATSIT');
  OkEtat:='X';
  end;
{$ENDIF}

TobP.PutValue ('ETATSTANDARD_IMPMODELE', stEtatStandard + Modele);
TobP.PutValue ('OKETAT', OkEtat);
TobP.PutValue ('NBEXEMPLAIRE', NbCopies);
if GetInfoParPiece(NaturePiece,'GPP_ACTIONFINI') = 'IMP' then TobP.PutValue ('GP_VIVANTE', '-');
end;

/////// Affichage PrintDialog pour impression dans fichier autre que pdf (speccif CEGID)
{$IFDEF MISESOUSPLI}
Function TOF_EditDocDiff_Mul.AffichePrintDialog : boolean ;
var aPrintDialog : TPrintDialog;
   // aPrinter : TPrintrer;
Begin
//     aPrinter:=Printer;
     aPrintDialog:=TPrintDialog.Create(Application);
     aPrintDialog.Options:=[poPageNums, PoPrintToFile];
     aPrintDialog.PrintRange:=prAllPages ;
     aPrintDialog.PrintToFile:=PrintToFile ;
     aPrintDialog.FromPage:=1 ;
     aPrintDialog.ToPage:=-1 ;
     aPrintDialog.MaxPage:=-1 ;
     aPrintDialog.MinPage:=0 ;
     Result:=aPrintDialog.Execute;
     if Result then PrinterIndex:=Printer.PrinterIndex ;
     aPrintDialog.free;
     //aPrinter.free;
End;
{$ENDIF}

procedure TOF_EditDocDiff_Mul.AlimPieceVivante;
BEGIN
inherited;
{$IFDEF AFFAIRE} // DBR Rassemblement
if ctxAffaire in V_PGI.PGIContexte then
    TCheckBox(GetControl('GP_VIVANTE')).State := cbGrayed;
{$ENDIF} // AFFAIRE
END;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLEditDocDiff_mul(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_EditDocDiff_Mul) then TOF_EditDocDiff_Mul(MaTOF).LanceEdition else exit;
end;

procedure AGLAlimPieceVivante( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_EditDocDiff_Mul) then TOF_EditDocDiff_Mul(MaTOF).AlimPieceVivante else exit;
end;

Initialization
registerclasses([TOF_EditDocDiff_Mul]);
RegisterAglProc('EditDocDiff_mul',TRUE,1,AGLEditDocDiff_mul);
RegisterAglProc( 'AlimPieceVivante',True,0,AGLAlimPieceVivante);
end.



