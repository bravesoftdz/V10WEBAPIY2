unit UTofAfPieceEdit_mul;

interface
uses
{$IFDEF EAGLCLIENT}
   eMul,Maineagl,eFiche,HPdfPrev ,UtileAGL,
{$ELSE}
   Mul,Fe_Main,Fiche,db,  dbTables,HDB,DBGrids,
//{$IFDEF V530}
//    EdtDoc,EdtEtat,
//{$ELSE}

//$ENDIF}
    EdtRDOC, EdtREtat,
{$ENDIF}
       URecupSQLModele,
      StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,M3VM , M3FP, Hqry, EntGC,
      FactTOB, // PL le 14/08/03 : suite modifs JLD facturation
      FactComm, FactUtil, UtofAFBASEPIECE_MUL,dicoaf,UtilMulTrt;

Type

     TOF_AfPieceEdit_mul = Class (TOF_AFBASEPIECE_MUL)
      private
        titre : string;
        TobComplete : TOB;
        procedure OnArgument (stArgument : string); override;
        procedure LanceEdition(zaction : string) ;
        procedure AlimPieceVivante;
        procedure EditeLesDocs (TobPiece : TOB;zaction : string);
        function RenseignePiece (stPrefixe, stTable, stEtatStandard : string;
                                     iInd : integer; TobP : TOB) : boolean;
        function ChercheCriteresDeTri : string;
    end ;

procedure AGLAfPieceEdit_mul(parms:array of variant; nb: integer ) ;

const
// libellés des messages
TexteMessage: array[1..3] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'Changement d''imprimé.' + #13 + 'Les documents suivants s''éditeront d''après le modèle '
          {3} ,'Une seule facture doit être sélectionnée pour l''apercu'
              );

implementation

procedure TOF_AfPieceEdit_mul.OnArgument (stArgument : string) ;
begin
Inherited;
titre := Ecran.caption;
SetControlChecked('GP_EDITEE',false);
end;

procedure TOF_AfPieceEdit_mul.LanceEdition(zaction : string) ;
var  St :string;
     //i,nbfac :integer;
     //Q:TQuery;
     req : string;
begin
St:= 'Confirmez vous l''édition de ces Factures';
If (PGIAskAf(st,titre)<> mrYes) then exit;
// on crée une TOB de toutes les pieces sélectionnées
TobComplete := TOB.Create ('Liste des pieces',NIL, -1);
// PA le 29/08/2001 - Fonction de traitement des enreg du mul externalisée
// mcd 19/08/02 Req := 'SELECT P.* FROM PIECE AS P LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE';
Req := 'SELECT P.* FROM PIECE P LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE';
TraiteEnregMulTable (TFMul(Ecran),Req,'GP_NATUREPIECEG;GP_SOUCHE;GP_NUMERO;GP_INDICEG','PIECE', 'GP_NUMERO','AFPIECEAFFAIRE',TobComplete, True);

EditeLesDocs (TObCOmplete,zaction);
TobComplete.free;
end;


procedure TOF_AfPieceEdit_mul.EditeLesDocs (TobPiece : TOB;zaction : string) ;
var iInd, iNbExemplaire,zexemp : integer;
    TobDet : TOB;
    CleDoc : R_CleDoc;
    stSql, stCle, stModele, sTri : string;
    stModeleSuivant,st,Bdupli : string;
    TL : TList;
    TT : TStrings;
    Bap,bdup,Baper : boolean;
    // mcd 29/01/02  ret : boolean;
    ChoixImp : boolean;  //mcd 29/01/02
    Pages : TpageControl;
BEGIN

Pages := nil;
bap := TCheckBox(GetControl('BAPERCU')).Checked;
// mcd 29/01/02 bdup := TCheckBox(GetControl('BDUPLICATA')).Checked;
   //mcd 29/01/02
Case TCheckBox(GetControl('BDUPLICATA')).State of
    cbGrayed   : Bdupli := '|' ;
    cbChecked  : Bdupli := 'X' ;
    cbUnChecked: Bdupli := '-' ;
    End;
//bdupli := TCheckBox(GetControl('BDUPLICATA')).state;  //mcd 29/01/02
TobPiece.Detail[0].AddChampSup ('ETATSTANDARD_IMPMODELE', True);
TobPiece.Detail[0].AddChampSup ('NBEXEMPLAIRE', true);
TobPiece.Detail[0].AddChampSup ('OKETAT', True);
for iInd := 0 to TobPiece.Detail.count - 1 do
    begin
      TobDet := TobPiece.Detail[iInd];
      TobDet.putValue('ETATSTANDARD_IMPMODELE','');
      TobDet.putValue('NBEXEMPLAIRE','0');


       // recup du modèle Client/type de piéce
      RenseignePiece ('GTP_', 'TIERSPIECE', '2', iInd, TobDet);
     // recup du modèle Parpièce (sauf si on a un modéle client/FAC et de l'action fini
        // mcd 12/07/02 appel pièce déplacer pour être fait après spécif client
        // sinon, on ne prenait plus en compte le modèle Client
        // plus suppression modif PL 03/09/01 qui traiteir le nb exemplaire : traiter dans
        // renseignePiece
      RenseignePiece ('GPP_', 'PARPIECE', '1', iInd, TobDet);

      if (TobDet.GetValue ('ETATSTANDARD_IMPMODELE') = '') then
        TobPiece.Detail[iInd].PutValue ('ETATSTANDARD_IMPMODELE', '0' + VH_GC.GCImpModeleDefaut);
      if (TobDet.GetValue ('NBEXEMPLAIRE') = '0') then
        TobPiece.Detail[iInd].PutValue ('NBEXEMPLAIRE', '1');

      if (THValComboBox(GetControl('MODELEEDITION')).value <> '') then
         Begin
         TobPiece.Detail[iInd].PutValue ('NBEXEMPLAIRE', '1');
         TobPiece.Detail[iInd].PutValue ('ETATSTANDARD_IMPMODELE', '0' + THValComboBox(GetControl('MODELEEDITION')).value);
         if (TRadioButton(GetControl('RBEDITETAT')).Checked = true) then
      	      TobPiece.Detail[iInd].PutValue ('OKETAT', 'X')
         else
      	       TobPiece.Detail[iInd].PutValue ('OKETAT', '-');
         End;


    end;
//tri par modéle   0:modéle par défaut, 1 :modéle CLient/FAC et 2 modéle du parpiéce
//  pourquoi ce tri sur l'origine du modèle ?

    sTri := ChercheCriteresDeTri; //

    TobPiece.Detail.Sort ('ETATSTANDARD_IMPMODELE'+ sTri);
    ChoixImp:=True;      // mcd 29/01/02
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


  if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then
    begin
    stSql:=RecupSQLModele('E','GPJ', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False));
    Pages := TPageControl.Create(Application);
    if (pos('ORDER BY',uppercase(stSql))<=0) then stSql:=stSql+' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE' ;
    end
  else
    begin
    stSql:=RecupSQLModele('L','GPI', stModele,'','','',' WHERE '+WherePiece(CleDoc,ttdPiece,False));
    if (pos('ORDER BY',AnsiUppercase(stSQL))<=0) then stSQL  := stSQL + ' ORDER BY GL_NUMLIGNE';
    end;

    Baper := Bap;
    zexemp := StrToInt (TobPiece.Detail[iInd].GetValue ('NBEXEMPLAIRE'));
    for iNbExemplaire := 1 to zexemp do
    begin
        TL := TList.Create ;
        TT := TStringList.Create ;
        TT.Add (stSql);
        TL.Add (TT);
               //mcd 29/01/02

        If (Bdupli='X') then Bdup:=True
          else if (Bdupli='-') then Bdup:=False
          else if InbExemplaire=1 then Bdup:=False
               else Bdup:=true;


        if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then
            Begin
       	 // mcd 29/01/02 V_PGI.NoPrintDialog:=True;   // 10/4/01
            V_PGI.NoPrintDialog:=Not(ChoixImp);  //mcd 29/01/02
            LanceEtat('E','GPJ',stModele,Baper, false, false,Nil,trim (stSql),'',BDup);
            End
        else
               //mcd 29/01/02 mise choiximp au lieu de False ==> demande l'impimante que la 1ere fois
            LanceDocument ('L', 'GPI', stModele, TL, Nil, Baper,ChoixImp, BDup);
        // mcd 29/01/02 ?? pas utiliser après ???  ret := V_PGI.QRPrinted;
        ChoixImp:=False;      // mcd 29/01/02
        BAper:=false;         // si plusieurs exemplaires pour une même facture, apercu uniquement sur la 1ere, ensuite on imprime directement les autres ...
        TT.Free;
        TL.Free;                                            
    end;

   if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then Pages.Free;
    if V_PGI.QRPrinted then
       TobPiece.Detail[iInd].PutValue ('GP_EDITEE', 'X');

    if (TCheckBox(Ecran.FindComponent ('BINFORMECHGTETAT')).Checked = True) and
       (iInd < TobPiece.Detail.Count - 1) then
        begin
        stModeleSuivant := copy (TobPiece.Detail[iInd + 1].GetValue ('ETATSTANDARD_IMPMODELE'), 2,
                                 length (TobPiece.Detail[iInd].GetValue ('ETATSTANDARD_IMPMODELE')) - 1);
        if stModele <> stModeleSuivant then
            begin
                 // mcd 12/07/01 ajout appel bonne table si etat et pas document
            if TobPiece.Detail[iInd].GetValue ('OKETAT') = 'X' then st := format(TexteMessage[2]+ ' : %s', [RechDom ('GCIMPETAT', stModeleSuivant, False)])
            else st := format(TexteMessage[2]+ ' : %s', [RechDom ('GCIMPMODELE', stModeleSuivant, False)]);
            If (PGIAskAf(st,titre)<> mrYes) then break;
           end;
        end;
    end;

    TobPiece.InsertorUpdateDB(false);

END;

function TOF_AfPieceEdit_mul.ChercheCriteresDeTri : string;
var
TriEdit1,TriEdit2,TriEdit3:THValComboBox;
begin
Result := '';
TriEdit1:=THValComboBox(GetControl('TRIEDITION1'));
TriEdit2:=THValComboBox(GetControl('TRIEDITION2'));
TriEdit3:=THValComboBox(GetControl('TRIEDITION3'));
if (TriEdit1<>nil) then
    if (TriEdit1.Value<>'') and (TriEdit1.Value<>'<<Aucun>>') then
    	Result:=Result+';'+TriEdit1.Value;

if (TriEdit2<>nil) then
    if (TriEdit2.Value<>'') and (TriEdit2.Value<>'<<Aucun>>') then
        	Result:=Result+';'+TriEdit2.Value;

if (TriEdit3<>nil) then
    if (TriEdit3.Value<>'') and (TriEdit3.Value<>'<<Aucun>>') then
    	Result:=Result+';'+TriEdit3.Value;


if (pos('GP_NUMERO', Result)=0) then
    Result := Result+';GP_NUMERO';
end;


//
//   Recherche du modéle , nombre d'exemplaire  dans Parpiece,et Tiers Piéce
//     et action fini  (slt dans Parpiéce)
//
function TOF_AfPieceEdit_mul.RenseignePiece (stPrefixe, stTable, stEtatStandard : string;
                                             iInd : integer; TobP : TOB) : boolean;
var TSql : TQuery;
    stSuff, stSelect, Modele  : string;
begin
Result := True;

if TobP.GetValue('GP_SAISIECONTRE') = 'X' then stSuff := 'CON'
else stSuff := '';

stSelect := 'SELECT ' + stPrefixe + 'IMPMODELE' + stSuff + ', ' +
                stPrefixe + 'IMPETAT' + stSuff + ', ' + stPrefixe + 'NBEXEMPLAIRE ';
if stPrefixe = 'GTP_' then
    begin
    stSelect := stSelect + 'FROM ' + stTable + ' WHERE ' + stPrefixe + 'TIERS="' +
                    TobP.GetValue ('GP_TIERS') + '" AND ';
    end else
    begin
    stSelect := stSelect + ', ' + stPrefixe + 'ACTIONFINI FROM ' + stTable + ' WHERE ';
   end;
stSelect := stSelect + stPrefixe + 'NATUREPIECEG="' + TobP.GetValue ('GP_NATUREPIECEG') + '"';

TSql := OpenSQL (stSelect, True);
if not TSql.Eof then
  begin
    // mcd 17/12/01 ajout pour prendre en compte les perso éventuelle par établissement
    {Document Nature/Etablissement}
  if stTable='PARPIECE' then
    begin    // on regarde d'abo si un état est personnalisé
    Modele := GetInfoParPieceCompl(TobP.GetValue ('GP_NATUREPIECEG'),TobP.GetValue ('GP_ETABLISSEMENT'),'GPC_IMPETAT'+StSuff) ;
    if (Modele = '') then
     begin   // sinon on regarde si on a un document
     Modele := GetInfoParPieceCompl(TobP.GetValue ('GP_NATUREPIECEG'),TobP.GetValue ('GP_ETABLISSEMENT'),'GPC_IMPMODELE'+StSuff) ;
     if (Modele <> '') then
       begin
       TobP.PutValue ('ETATSTANDARD_IMPMODELE',stEtatStandard +Modele);
       TobP.PutValue ('OKETAT', '-');
       end;
     end
    else  begin
     TobP.PutValue ('OKETAT', 'X');
     TobP.PutValue ('ETATSTANDARD_IMPMODELE',stEtatStandard +Modele);
     end;
    end;
       // pas de récupération du modèle lu si déjà chargé (cas existe pour le client si lecture piece ...)
  if (TobP.GetValue('ETATSTANDARD_IMPMODELE') = '') then
    Begin 
    if TSql.Fields[1].AsString <> '' then
      begin
      TobP.PutValue ('ETATSTANDARD_IMPMODELE', stEtatStandard + TSql.Fields[1].AsString);
      TobP.PutValue ('OKETAT', 'X');
      end
    else
      begin
      if TSql.Fields[0].AsString <> '' then
        Begin
        TobP.PutValue ('ETATSTANDARD_IMPMODELE', stEtatStandard + TSql.Fields[0].AsString);
        TobP.PutValue ('OKETAT', '-');
        end
      end;
    End;

    // Modif PL le 03/09/01
    // mcd 12/07/02 ajout test pour ne pas écraser l enbre exemplaire si déjà renseigné sur client ou pièce
if (TobP.GetValue ('NBEXEMPLAIRE') =0) then begin
 if (TSql.Fields[2].AsString <> '') and (TSql.Fields[2].AsString <> '0') then
    TobP.PutValue ('NBEXEMPLAIRE', TSql.Fields[2].AsString);
 end;
    // Fin Modif PL le 03/09/01

   if (stPrefixe = 'GPP_') and (TSql.Fields[3].AsString = 'IMP') and
       (TobP.GetValue ('GP_VIVANTE') = 'X') then TobP.PutValue ('GP_VIVANTE', '-');

  end
    else
        Result := False;

Ferme (TSql);
end;


procedure TOF_AfPieceEdit_mul.AlimPieceVivante;
BEGIN
inherited;
TCheckBox(GetControl('GP_VIVANTE')).State := cbGrayed;
END;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLAfPieceEdit_mul(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AfPieceEdit_mul) then TOF_AfPieceEdit_mul(MaTOF).LanceEdition(Parms[1]) else exit;
end;

procedure AGLAlimPieceVivante( parms: array of variant; nb: integer );
var  F : TForm ;
     MaTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFMul) then MaTOF:=TFMul(F).LaTOF else exit;
if (MaTOF is TOF_AfPieceEdit_mul) then TOF_AfPieceEdit_mul(MaTOF).AlimPieceVivante else exit;
end;


Initialization
registerclasses([TOF_AfPieceEdit_mul]);
RegisterAglProc('AfPieceEdit_mul',TRUE,1,AGLAfPieceEdit_mul);
RegisterAglProc( 'AlimPieceVivante',True,0,AGLAlimPieceVivante);
end.
