unit UTofGcModifLot;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,GCMzsUtil,UtilArticle,
{$IFDEF EAGLCLIENT}
      Maineagl,eMul,
{$ELSE}
      Fiche,mul, DBGrids, Fe_Main,HDB,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      M3VM , M3FP,utilGC,AglInitGc,UtilPgi
       ;

Type

     TOF_GcModifLot = Class (TOF)
        // Modif BTP
        CCV : THValComboBox ;
        TypeNomen : THValComboBox;
        procedure OnArgument (S : String ) ; override ;
        // -- Fin Modif BTP

        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
// Ancienne version       procedure MiseAJourDuLot(L:THDBGrid;Q:TQuery);
        procedure ModificationParLot ;
       // procedure ModificationArticle(Cle,NomTable:string ; TobRecupChamp: TOB) ;
       // Modif BTP
       private
{$IFDEF BTP}
       procedure ChangeTypeArt(Sender: Tobject);
    procedure RenommageEnteteColonnes;
{$ENDIF}
       // -----
     END ;

     TOF_CompteRendu = Class (TOF)
        procedure OnLoad; override ;

     END;

procedure AGLModificationParLot(parms:array of variant; nb: integer ) ;
//procedure CopieTob(source,destination:TOB) ;
procedure MessageErreur(Article,CodeArticle, LibelleArticle, MsgErreur :string; TobMere :TOB; CodeErreur: integer);


implementation

uses Grids;
// Modif BTP
procedure TOF_GcModifLot.OnArgument (S : String );
begin
{$IFDEF BTP}
CCV := THValComboBox (ecran.FindComponent('GA_TYPEARTICLE'));
CCV.plus := ' AND CO_CODE IN ("ARP","MAR","OUV","POU","PRE","FRA")';
CCV.Vide := false;
CCV.ReLoad;
CCV.value := 'MAR';
CCV.OnChange := ChangeTypeArt;
TypeNomen := THValComboBox (ecran.findComponent('GA_TYPENOMENC'));
{$ENDIF}
MajChampsLibresArticle(TForm(Ecran));
SetControlVisible('BSELECTALL',True); //JS 17/04/03
end;

{$IFDEF BTP}
procedure TOF_GcModifLot.ChangeTypeArt ( Sender : Tobject);
var icol : integer;
    CC : THlabel;
begin
  if (CCV.Value = 'OUV') or (CCV.value = 'NOM')   then
  begin
    TypeNomen.visible := true;
    TypeNomen.value := 'OUV';

    TcheckBox(GetControl ('GA_TENUESTOCK')).Visible := false;
    TCheckBox(GetControl ('GA_TENUESTOCK')).State  := cbGrayed;

    GetControl ('TGA_TYPENOMENC').Visible := true;
  	TFMul(ecran).SetDBListe ('BTMULOUVRAGE');
    THValComboBox(getControl('GA_FAMILLENIV1')).DataType := 'BTFAMILLEOUV1';
    THValComboBox(getControl('GA_FAMILLENIV2')).DataType := 'BTFAMILLEOUV2';
    THValComboBox(getControl('GA_FAMILLENIV3')).DataType := 'BTFAMILLEOUV3';
  end else
  begin
    if CCV.value = 'PRE' then  TFMul(ecran).SetDBListe ('BTPRESTATIONSMUL')
                         else  TFMul(ecran).SetDBListe ('BTRECHARTICLE');
    GetControl ('GA_TENUESTOCK').Visible := true;
    TypeNomen.visible := false;
    TypeNomen.value:='';
    GetControl ('TGA_TYPENOMENC').Visible := false;
    THValComboBox(getControl('GA_FAMILLENIV1')).DataType := 'GCFAMILLENIV1';
    THValComboBox(getControl('GA_FAMILLENIV2')).DataType := 'GCFAMILLENIV2';
    THValComboBox(getControl('GA_FAMILLENIV3')).DataType := 'GCFAMILLENIV3';
  end;

  for iCol:=1 to 3 do
  begin
    CC:=THLabel(GetCOntrol('TGA_FAMILLENIV'+InttoStr(iCol)));
    if copy(ecran.name,1,13) <>'BTARTICLE_MUL' then
    begin
      if ThvalComboBox(GetControl('GA_TYPEARTICLE')).Value = 'OUV' then
      begin
        CC.Caption:=RechDom('BTLIBOUVRAGE','BO'+InttoStr(iCol),false);
      end else
      begin
        CC.Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(iCol),false);
      end;
    end else
    begin
    end;
  end;

end;
{$ENDIF}

procedure TOF_GcModifLot.RenommageEnteteColonnes;
var i : integer;
		Gr : THDbgrid;
    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName);
    if copy(UpperCase (stChamp),1,6)='LIBPRE' then
    begin
      libelle := RechDom('GCLIBFAMILLE','LF'+Copy(stChamp,7,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,11)='GA_LIBREART' then
    begin
      libelle := RechDomZoneLibre ('AT'+Copy(stChamp,12,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,12)='GA_DATELIBRE' then
    begin
      libelle := RechDomZoneLibre ('AD'+Copy(stChamp,13,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,11)='GA_VALLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AM'+Copy(stChamp,12,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,12)='GA_CHARLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AC'+Copy(stChamp,13,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,12)='GA_BOOLLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AB'+Copy(stChamp,13,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end;


  end;
end;

// -------------

procedure TOF_GcModifLot.OnLoad;
begin
Ecran.Caption:='Saisie en lot d''article';
RenommageEnteteColonnes;
end;


procedure TOF_GcModifLot.OnUpdate;
var  F : TFMul ;
    i:integer;
    CC:THLabel ;
begin
inherited ;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;

for i:=1 to 3 do
    begin
    CC:= THLabel(TFMul(F).FindComponent('TGA_FAMILLENIV'+InttoStr(i)));
    CC.Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(i),false);
    end;
{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
{$ENDIF}
RenommageEnteteColonnes;
end;

/////////////// ModificationParLotDesTiers //////////////
procedure TOF_GcModifLot.ModificationParLot;
Var F : TFMul ;
    ArticleSpecif: String ;
    Parametrages : String ;
    TheModifLot : TO_ModifParLot;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin MessageAlerte('Aucun élément sélectionné'); exit; end;
//if HShowMessage('0;Confirmation;Confirmez vous la modification des clients?;Q;YN;N;N;','','')<>mrYes then exit;
TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.Titre := 'Modification par lot des Articles';
TheModifLot.FCode := 'GA_ARTICLE';
TheModifLot.TableName := 'ARTICLE';
//
// Modif BTP
//
{$IFDEF BTP}
   TheModifLot.Nature := 'BTP';
   if (CCV.value = 'NOM') or (CCV.value = 'OUV') then TheModifLot.StContexte := TypeNomen.Value
                                                 else TheModifLot.StContexte := CCV.Value;
   if CCV.Value = 'POU' Then TheModifLot.FicheAOuvrir := 'BTARTPOURCENT'
                        else if (CCV.Value = 'PRE') then
                                TheModifLot.FicheAOuvrir := 'BTPRESTATION'
                        else TheModifLot.FicheAOuvrir := 'BTARTICLE';
{$ELSE}
{$IFDEF GPAO} { GPAO0 }
   TheModifLot.Nature := NatureFicheArticle;
   TheModifLot.FicheAOuvrir := NomFicheArticle;
{$ELSE}
// Nom de la fiche à ouvrir en fontion de V_PGI.PGIContexte
ArticleSpecif:=IsArticleSpecif('MAR') ;
if ArticleSpecif='FicheAffaire' then
   begin
   TheModifLot.Nature := 'AFF';
   TheModifLot.FicheAOuvrir := 'AFARTICLE';
   end
else if ArticleSpecif='FicheModeArt' then
   begin
   TheModifLot.Nature := 'MBO';
   TheModifLot.FicheAOuvrir := 'ARTICLE';
   end
else
   begin
   TheModifLot.Nature := 'GC';
   TheModifLot.FicheAOuvrir := 'GCARTICLE';
   end;
{$ENDIF}
{$ENDIF}

ModifieEnSerie(TheModifLot, Parametrages) ;

end;


{  ANNCIENNE VERSION
////////////// Modification d'un article //////////////////////////

procedure TOF_GcModifLot.ModificationArticle(Cle,NomTable:string ; TobRecupChamp: TOB) ;
var TOM_Modif :TOM;
    TOB_Modif :TOB;
Begin
LastError:=0;
LastErrorMsg:='';
TOB_Modif:=TOB.Create(NomTable, NIL, -1) ;
TOM_Modif :=CreateTOM(NomTable,Nil,False,False);
TOB_Modif.SelectDB('"'+Cle+'"',nil);
CopieTob(TobRecupChamp,TOB_Modif);
  if TOM_Modif.VerifTOB(TOB_Modif) then
     TOB_Modif.UpdateDB
  Else
    begin
    LastError:= TOM_Modif.LastError;
    LastErrorMsg:= TOM_Modif.LastErrorMsg;
    end;

TOM_Modif.Free;
TOB_Modif.Free;
end;


//////////////////////// Mise à jour du lot d'article ///////////////////////////
procedure TOF_GcModifLot.MiseAJourDuLot(L:THDBGrid;Q:TQuery);
var i  : integer;
Article ,CodeArticle ,Table, Libelle : string;
TobErreur :TOB;
begin
if HShowMessage('0;Confirmation;Confirmez vous la modification des données?;Q;YN;N;N;','','')<>mrYes then exit;
TobErreur:= TOB.Create('les erreurs',NIL,-1);
if L.AllSelected then
   BEGIN
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
     BEGIN
     Article:=TFmul(Ecran).Q.FindField('GA_ARTICLE').asstring ;
     CodeArticle:=TFmul(Ecran).Q.FindField('GA_CODEARTICLE').asstring ;
     Libelle:=TFmul(Ecran).Q.FindField('GA_LIBELLE').asstring ;
     Table:=GetTableNameFromDataSet(Q);
     MoveCur(False);
     ModificationArticle(Article,Table,LaTOB);
     if LastError<>0 then
       BEGIN
       MessageErreur(Article,CodeArticle,Libelle, LastErrorMsg, TobErreur, LastError);
       END;
     Q.NEXT;
     END;
     L.AllSelected:=False;
   END ELSE
   BEGIN
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
     BEGIN
     L.GotoLeBOOKMARK(i);
     Article:=TFmul(Ecran).Q.FindField('GA_ARTICLE').asstring ;
     CodeArticle:=TFmul(Ecran).Q.FindField('GA_CODEARTICLE').asstring ;
     Libelle:=TFmul(Ecran).Q.FindField('GA_LIBELLE').asstring ;
     Table:=GetTableNameFromDataSet(Q);
     MoveCur(False);
     ModificationArticle(Article,Table,LaTOB);
     if LastError<>0 then
       BEGIN
       MessageErreur(Article,CodeArticle,Libelle, LastErrorMsg, TobErreur, LastError);
       END;
     END;
   L.ClearSelected;
   END;
FiniMove;
If LastError <> 0 then
   begin
   TheTob:=TobErreur;
   AGLLanceFiche('GC','GCAFFICHEERREUR','','','') ;
   end
   else
   HShowMessage('2;Information;Tous les articles ont été modifiés;I;O;O;N;','','');

TobErreur.Free;
End;


//////////// Modification par lot des articles ///////////////
procedure TOF_GcModifLot.ModificationParLot ;
var  F : TFMul ;
     NouveauChamp, St:string;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   MessageAlerte('Aucun élément sélectionné');
   exit;
   end;
LaTOB:=TOB.Create('_modiflot',NIL,-1);
St:=GetControlText('GCCHAMPARTMODIF');
While St <> '' do
      begin
      NouveauChamp:=uppercase(Trim(ReadTokenSt(St)));
      LaTOB.AddChampSup(NouveauChamp,TRUE);
      end;
TheTOB:=LaTOB;
AGLlancefiche('GC','GCARTICLE','','','MODIFLOT;ACTION=CREATION') ;
if TheTOB<>Nil then
   begin
   LaTOB:=TheTOB;
   MiseAJourDuLot(F.FListe, F.Q);
   TheTOB:=NIL;
   LaTOB.Free;
   end;
end;
}

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLModificationParLot(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_GcModifLot) then TOF_GcModifLot(TOTOF).ModificationParLot else exit;
end;

{
/////////////// Copie de Tob ////////////////////
procedure CopieTob(source,destination:TOB) ;
var i:integer;
NomChamp: string;
ValeurChamp: Variant ;
Begin
    For i:=0 to source.ChampsSup.count-1 do
    begin
    NomChamp:= TCS(source.ChampsSup[i]).Nom ;
    ValeurChamp:= source.GetValue(NomChamp) ;
    destination.PutValue(NomChamp,ValeurChamp) ;
    end;
end;
}

//////////// Création de tob fille pour le THGrid des erreurs /////////
procedure MessageErreur(Article,CodeArticle, LibelleArticle, MsgErreur :string; TobMere :TOB; CodeErreur: integer);
var NomTobFille :TOB ;
begin
NomTobFille:= Tob.Create('une erreur',TobMere,-1) ;
NomTobFille.AddChampSup('GA_ARTICLE',False);
NomTobFille.AddChampSup('GA_CODEARTICLE',False);
NomTobFille.AddChampSup('GA_LIBELLE',False);
NomTobFille.AddChampSup('_CodeErreur',False);
NomTobFille.AddChampSup('_MsgErreur',False);
NomTobFille.PutValue('GA_ARTICLE',Article);
NomTobFille.PutValue('GA_CODEARTICLE',CodeArticle);
NomTobFille.PutValue('GA_LIBELLE',LibelleArticle);
NomTobFille.PutValue('_CodeErreur',CodeErreur);
NomTobFille.PutValue('_MsgErreur',MsgErreur);
end;



////////////// TOF CompteRendu : THGrid pour affichage des erreurs /////////////
procedure TOF_CompteRendu.OnLoad ;
Var Messages: TControl;
GridMessage: THGrid;
begin
Ecran.Caption:=' Liste des erreurs rencontrées ';
Messages:=GetControl('GRIGMESSAGEERREUR') ;
GridMessage:=THGrid(Messages) ;
GridMessage.ColWidths[0]:=0 ;
if LaTob <> NIL then
LaTob.PutGridDetail(GridMessage,False,False,'',False) ;
LaTob:=Nil;
end;


Initialization
registerclasses([TOF_GcModifLot]);
registerclasses([TOF_CompteRendu]);
RegisterAglProc('ModificationParLot',TRUE,2,AGLModificationParLot);
end.



