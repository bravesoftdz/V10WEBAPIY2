unit UTofAfPieceProAnu_Mul;
{*************************************************************************
  Annulation des documents PROVISOIRES
   Principe Multicritère avec Multi-Sélection Possible.
   En sortie de Mul:
   * je récupère les factures à annuler  :
   Suppresion des factures , maj de l'activité,


**************************************************************************}
interface                                 
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,EntGC,
{$IFDEF EAGLCLIENT}
    eMul, MainEAGL,
{$ELSE}
   Mul,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,DBGrids,HDB, FE_Main,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF,AffaireUtil,Dicobtp,SaisUtil,M3FP,UTofAfBasePiece_Mul,
      Hstatus, Utob, FactGrp,Factcomm,factutil,UtilMulTrt,ParamSoc, FactTOB,utilpgi,uEntCommun,UtilTOBPiece;

       procedure AGLAnnulationFacture( Parms : array of variant ; nb : integer );
Type
     TOF_AFPIECEPROANU_MUL = Class (TOF_AFBASEPIECE_MUL)
     public

     		procedure OnArgument(stArgument : String ) ; override ;
        //Procédures
        procedure AnnulApprec(Tobpiece : Tob);
        procedure AnnulationFacture ;
        procedure AnnulConso(TOBPiece: Tob);
        procedure ControleChamp(Champ, Valeur: String);
        procedure ControleCritere(Critere: String);
        procedure DeflagerFactureOrigine(pTobPiece : Tob);
        procedure SupFactPro(Tobpiece : Tob);
        procedure TrtAnuFactPro;
        procedure DeflagueLesEcheanceslignes(TOBpiece : TOB);

     private

     //Variables Globales
     Tobcomplete 	: TOB;
     Titre 				: string;
     Statut				: String;

     //procedure bSelectAff1Click(Sender: TObject);     override ;
     procedure AnnulFactureAffaire(TOBPiece: Tob);

     END ;

procedure AFLanceFiche_Mul_Annu_PiecePro(Argument: string);

implementation

procedure TOF_AFPIECEPROANU_MUL.OnArgument(stArgument : String );
var Critere	: String;
		Champ		: String;
    Valeur  : String;
    x				: integer;

Begin
	fMulDeTraitement := true;
Inherited;
	FTableName := 'PIECE';
	// Recup des critères
	titre := Ecran.caption;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(StArgument)));

	While (Critere <>'') do
  	  Begin
    	if Critere<>'' then
      	 Begin
         X:=pos(':',Critere);
       	 if x=0 then X:=pos('=',Critere);
         if x<>0 then
            begin
            Champ :=copy(Critere,1,X-1);
            Valeur:=Copy (Critere,X+1,length(Critere)-X);
            End
         else
           Champ := Critere;
         ControleChamp(Champ, Valeur);
         Critere:=(Trim(ReadTokenSt(stArgument)));
    		 End;
      End;

  UpdateCaption(Ecran);

End;

Procedure TOF_AFPIECEPROANU_MUL.ControleChamp(Champ : String;Valeur : String);
Begin

  if Champ = 'STATUT' then
     Begin
    if Valeur = 'APP' then
        Begin
      Statut := 'W';
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', Statut)
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', Statut);
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel');
        SetControlText('TGP_AFFAIRE', 'Appel');
        end
     Else if valeur = 'INT' then
          Begin
      Statut := 'I';
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', Statut)
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', Statut);
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="' + Statut + '" AND GP_GENERAUTO="CON"');
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Contrat');
          SetControlText('TGP_AFFAIRE', 'Contrat');
          end
     Else if valeur = 'AFF' then
          Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', Statut)
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', Statut);
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "")');
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Chantier');
          SetControlText('TGP_AFFAIRE', 'Chantier');
          end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('XX_WHERE', ' AND AFF_AFFAIRE0 IN ("W","A")')
      Else if assigned(GetControl('AFFAIRE0')) then SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)  IN ("W","A")');
      SetCOntrolText ('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A","W")');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire de Regroupement');
      SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Affaire de Regroupement');
      SetControlText('TGP_AFFAIRE', 'Affaire de Regroupement');
    end
     Else if valeur = 'PRO' then
          Begin
      Statut := 'P';
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', Statut)
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', Statut);
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel d''offre');
          SetControlText('TGP_AFFAIRE', 'Appel d''offre');
          end
     else
          Begin
      Statut := '';
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', Statut)
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', Statut);
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Affaire');
          SetControlText('TGP_AFFAIRE', 'Affaire');
    end;
    //
  end;

end;

Procedure TOF_AFPIECEPROANU_MUL.ControleCritere(Critere : String);
Begin
end;

procedure TOF_AFPIECEPROANU_MUL.AnnulationFacture ;
var  St :string;
     //i :integer;
     //io   : TIoErr ;
     Req : string;
begin

  if not VH_GC.GCIfDefCEGID then
     if Blocage (['nrGener'],True,'nrGener') then  exit;

  St:= 'Confirmez vous l''annulation de ces Factures Provisoires';

  If (PGIAskAf(st,titre)<> mrYes) then exit;

  // on crée une TOB de toutes les pieces sélectionnées
  TobComplete := TOB.Create ('Liste des pieces',NIL, -1);

  // PA le 26/09/2001 - Fonction de traitement des enreg du mul externalisée
  // mcd 19/09/02 Req :=  'SELECT P.* FROM PIECE AS P LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE =GP_AFFAIRE';
  Req :=  'SELECT P.* FROM PIECE P LEFT OUTER JOIN AFFAIRE ON AFF_AFFAIRE = GP_AFFAIRE';
  TraiteEnregMulTable(TFMul(Ecran),Req,'GP_NATUREPIECEG;GP_SOUCHE;GP_NUMERO;GP_INDICEG','PIECE', 'GP_NUMERO','AFPIECEAFFAIRE',TobComplete, True);

  // on appelle la fct générique d'annulation des facture
  Transactions(TrtAnuFactPro,1) ;
  TobComplete.free;

  if not VH_GC.GCIfDefCEGID then
    Bloqueur ('nrGener',False);

end;

procedure TOF_AFPIECEPROANU_MUL.DeflagerFactureOrigine(pTobpiece : Tob);
var
  vSt   : String;
  vTob  : Tob;
  vQr   : TQuery;
  i     : Integer;

begin

  if pTobpiece.GetValue('GP_CREEPAR') = 'REG' then
  begin
    //recherche des numeros des documents d'origine
    vSt := 'SELECT AFR_FORCODE, AFR_ONATUREPIECEG, AFR_OSOUCHE, AFR_ONUMERO, AFR_OINDICEG ';
    vSt := vSt + ' FROM AFREVISION, PIECE ';
    vSt := vSt + ' WHERE AFR_NATUREPIECEG = GP_NATUREPIECEG ';
    vSt := vSt + ' AND AFR_SOUCHE = GP_SOUCHE ';
    vSt := vSt + ' AND AFR_NUMERO = GP_NUMERO ';
    vSt := vSt + ' AND AFR_INDICEG = GP_INDICEG ';
    vSt := vSt + ' AND AFR_NATUREPIECEG ="' + pTobpiece.getValue('GP_NATUREPIECEG') + '"';
    vSt := vSt + ' AND AFR_SOUCHE ="' + pTobpiece.getValue('GP_SOUCHE') + '"';
    vSt := vSt + ' AND AFR_NUMERO =' + intToStr(pTobpiece.getValue('GP_NUMERO'));
    vSt := vSt + ' AND AFR_INDICEG =' + intToStr(pTobpiece.getValue('GP_INDICEG'));

    vTob := Tob.create('Mes revisions', nil, -1);
    vQr := nil;
    try
      vQR := OpenSql(vSt, True);
      if Not vQR.Eof then
        vTob.LoadDetailDB('AFREVISION', '', '', vQr, False) ;
                                
      // on parcours toutes les lignes de la piece annulée, pour deflaguer regularisées
      // les lignes des pieces d'origine
      for i := 0 to vTob.detail.count -1 do
      begin
        if (i = 0) or
           ((i <> 0) and
            (vTob.detail[i].getValue('AFR_ONUMERO') <>  vTob.detail[i-1].getValue('AFR_ONUMERO'))) then
        begin
          if vTob.detail[i].GetValue('AFR_FORCODE') <> '' then
          begin
            vSt := 'UPDATE LIGNE SET GL_REGULARISE = "-" ';
            vSt := vSt + ' WHERE GL_NATUREPIECEG ="' + vTob.detail[i].getValue('AFR_ONATUREPIECEG') + '"';
            vSt := vSt + ' AND GL_SOUCHE ="' + vTob.detail[i].getValue('AFR_OSOUCHE') + '"';
            vSt := vSt + ' AND GL_NUMERO =' + intToStr(vTob.detail[i].getValue('AFR_ONUMERO'));
            vSt := vSt + ' AND GL_INDICEG =' + intToStr(vTob.detail[i].getValue('AFR_OINDICEG'));
            ExecuteSql(vst);                    
          end;
        end;                        
      end;

    Finally
      if vQR <> nil then ferme(vQR);
      vTob.Free;
    End;
  end;
end;

procedure TOF_AFPIECEPROANU_MUL.TrtAnuFactPro;
var i : integer;
    TobPiece : TOB;
    nat:string;
    DEV : Rdevise;
BEGIN

  InitMove(TOBComplete.Detail.Count,'');

  for i:=0 to TOBComplete.Detail.Count-1 do
  BEGIN
    MoveCur(False);
    TOBPiece:=TOBComplete.Detail[i];
    if TOBpiece.getvalue('GP_AFFAIRE')='' then
    begin
			MajEcheance(nil, TOBpiece,'ANU') ;
    end;
    // revision de prix
    // deflager les factures d'origine de la facture de régul
    if GetParamSoc('SO_AFREVISIONPRIX') then
      DeflagerFactureOrigine(TOBPiece);

    //Réinitialisation Conso dans le cas d'un appel
    if (Statut='W') or (STATUT = 'APP') or (Statut = 'GRP') then AnnulConso(TobPiece);

	  SupFactPro(Tobpiece);

    // à ne pas faire dans le cas de l'avoir provisoire
    If (NatPiece<>nil)    then nat :=  natpiece.value; // mcd 09/10/02
    If (NatPieceMul<>nil)    then nat :=  natpieceMul.text; // mcd 09/10/02
    if (nat <> 'APR') then
    begin
      AnnulApprec(TobPiece);
      MajGestionAffaire(NIL,TobPiece,nil,nil,nil,'ANU',DEV);
    end;
  End;

  FiniMove;

END;

procedure TOF_AFPIECEPROANU_MUL.SupFactPro(Tobpiece : Tob);
var  numpiece			: string;
		 req 					: string;
     CodeAffaire	: string;
     QQ						: TQUERY;
     cledoc 			: R_CleDoc ;
     TypeAffaire : string;
     DevisExiste : boolean;
BEGIN

  numpiece :=  EncodeRefPiece(TOBPIECE);
  CodeAffaire := TOBPIECE.GetValue('GP_AFFAIRE');
  DecodeRefPiece(numpiece,Cledoc);
  ExecuteSQL('DELETE FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False)) ;
  ExecuteSQL('DELETE FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)) ;
  ExecuteSQL('DELETE FROM LIGNECOMPL WHERE '+WherePiece(CleDoc,ttdLigneCompl,False)) ;
  ExecuteSQL('DELETE FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False)) ;
  ExecuteSQL('DELETE FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False)) ;
  ExecuteSQL('DELETE FROM PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False)) ;
  ExecuteSQL ('DELETE FROM ACOMPTES WHERE '+wherePiece(Cledoc,ttdAcompte,False));//mcd 28/10/02
  ExecuteSQL ('DELETE FROM VENTANA WHERE YVA_IDENTIFIANT="GL" AND YVA_IDENTIFIANT ="'+EncodeRefCPGescom(TOBPiece)+'"');
  if GetParamSoc('SO_GCPIECEADRESSE') then
     BEGIN  //mcd 14/02/03
     ExecuteSQL('DELETE FROM PIECEADRESSE WHERE '+WherePiece(CleDoc,ttdPieceAdr,False)) ;
     END ;

  //Remise à jour du code état dans l'affaire initiale
  Req := 'SELECT AFF_STATUTAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAffaire + '"';
  QQ := OpenSql (Req,true);

  if not QQ.eof then
  begin
    TypeAffaire := QQ.FindField ('AFF_STATUTAFFAIRE') .AsString;
  end;
  ferme (QQ);

  if TypeAffaire = 'APP' then
  Begin
    DevisExiste := ExisteSQL ('SELECT GP_NUMERO FROM PIECE WHERE GP_AFFAIRE="'+ CodeAffaire + '" AND GP_NATUREPIECEG="DAP"');
    if DevisExiste then
    begin
      ExecuteSQL('UPDATE LIGNE SET GL_VIVANTE="X",GL_QTERESTE=GL_QTEFACT, GL_MTRESTE=GL_MONTANTHTDEV WHERE GL_AFFAIRE="'+ CodeAffaire + '" AND GL_NATUREPIECEG="DAP"');
      ExecuteSQL('UPDATE PIECE SET GP_VIVANTE="X" WHERE GP_AFFAIRE="'+ CodeAffaire + '" AND GP_NATUREPIECEG="DAP"');
      ExecuteSQL('UPDATE AFFAIRE SET AFF_ETATAFFAIRE="REA" WHERE AFF_AFFAIRE="'+ CodeAffaire + '"');
    end else
    begin
      ExecuteSQL('UPDATE AFFAIRE SET AFF_ETATAFFAIRE="REA" WHERE AFF_AFFAIRE="'+ CodeAffaire + '"');
    end;
  end;

  if GetParamSoc('SO_AFREVISIONPRIX') then
    ExecuteSQL('DELETE FROM AFREVISION WHERE '+ WherePiece(CleDoc,ttdRevision,False));

  if GetParamSoc('SO_AFVARIABLES') then
    ExecuteSQL('DELETE FROM AFORMULEVARQTE WHERE '+ WherePiece(CleDoc,ttdVariable,False));

END;

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 06/09/2001
Modifié le ... :   /  /
Description .. : Annulation facture.
Suite ........ : Gestion facture d'appréciation
Suite ........ :  * suppression des B/M générés
Suite ........ :  * détopgae des acompte en cours de reprise
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPIECEPROANU_MUL.AnnulApprec(Tobpiece : Tob);
var  numpiece,req,CodeAff,zgen : string;
      cledoc : R_CleDoc ;
      {nb ,}znum: Integer;
      QQ :Tquery;
BEGIN

  numpiece :=  EncodeRefPiece(TOBPIECE);
  DecodeRefPiece(numpiece,Cledoc);
  CodeAff := TobPiece.GetValue('GP_AFFAIRE');
  znum:=0;   // mcd 07/03/02

  // recup du n° de l'appréciation à partir du n° de Piece Provisoire
  Req := 'SELECT AFA_NUMECHE,AFA_GENERAUTO FROM FACTAFF Where AFA_AFFAIRE="'+CodeAff+'" AND AFA_TYPECHE="APP" '
  + ' and AFA_NUMPIECE = "'+numpiece+'"';        // Pas triés ?????

  // PL le 09/01/02 pour optimiser le nombre d'enregistrements lus en eagl
  QQ := nil;
  try
    QQ:=OpenSQL(Req,True,1);

  if Not QQ.EOF then
     begin
     QQ.First;  // classées par ordre décroissant
     znum := QQ.Fields[0].AsInteger;
     zgen := QQ.Fields[1].AsString;
     end;
  finally
    Ferme(QQ);
  end;

  if (znum = 0) then exit;

// suppresion des lignes de B/M s'il y en a
  Req := 'DELETE  FROM ACTIVITE WHERE ACT_TYPEACTIVITE = "BON" and ACT_AFFAIRE="'
  + CodeAff+'"' +' and ACT_NUMAPPREC ='+inttostr(znum);
  ExecuteSQL(req);

  Req := 'UPDATE PIECE SET GP_FACREPRISE=0 WHERE GP_AFFAIRE="' + CodeAff+'"' +' and GP_FACREPRISE ='+inttostr(znum);
  ExecuteSQL(req);

  // mcd 30/07/03 on ne seconetnet pas d'affacer l'enrgt d'appréciaiton, mais on l'efface
  req := 'delete  from FACTAFF ' + ' where afa_affaire="'+Codeaff+'"' + ' and AFA_NUMECHE ="'+IntToStr(Znum)+ '"';
  ExecuteSQL(req);

  //mcd 25/08/2003 il faut aussi effacer la date de la dernière situation
  Req := 'UPDATE AFFAIRE SET AFF_NUMSITUATION=0,AFF_DATESITUATION="'+UsDateTime(Idate1900)+'"';
  ExecuteSQL(Req);

END;

Procedure TOF_AFPIECEPROANU_MUL.AnnulConso(TOBPiece : Tob);
Var Req		: String;
		QQ		: TQuery;
		TobL	: TOB;
    I			: Integer;
		Num   : Integer;
    Naff	: String;
    NumLig: Integer;
Begin

	Num := TOBPiece.GetValue('GP_NUMERO');

  if Statut ='GRP' then
     AnnulFactureAffaire(TOBPiece)
  else
	   Naff := TOBPiece.GetValue('GP_AFFAIRE');

	TobL := Tob.create('LES LIGNES', nil, -1);

  Req := 'SELECT GL_NUMLIGNE FROM LIGNE ';
  Req := Req + 'WHERE GL_NATUREPIECEG="' + TOBPiece.getValue('GP_NATUREPIECEG') + '"';
  Req := Req + '  AND GL_SOUCHE="' + TOBPiece.getValue('GP_SOUCHE') + '"';
  Req := Req + '  AND GL_NUMERO =' + IntToStr(Num);
  Req := Req + '  AND GL_ARTICLE <> ""';

  TobL.LoadDetailDBFromSQL('LIGNE',req,false);

  if TOBL = nil then exit;

  //Remise en attente de facturation des lignes conso.
  For I := 0 To TobL.Detail.count -1 Do
  		Begin
      NumLig := TobL.Detail[i].GetValue('GL_NUMLIGNE');
		  Req := 'UPDATE CONSOMMATIONS SET BCO_FACTURABLE="A"';
	    Req := Req + ', BCO_NATUREPIECEG="", BCO_SOUCHE="", BCO_NUMERO=0, BCO_NUMORDRE=0';
	    Req := Req + ' WHERE BCO_NUMERO=' + IntToStr(Num) + ' AND BCO_NUMORDRE=' + IntToStr(NumLig);
      ExecuteSQL(Req);
      end;

end;

Procedure TOF_AFPIECEPROANU_MUL.AnnulFactureAffaire(TOBPiece : Tob);
var Req 	 : String;
    Naff	 : String;
    TOBGLC : Tob;
    I			 : Integer;
Begin

  TobGLC := Tob.create('LES COMPLEMENTS LIGNES', nil, -1);

	Req := 'SELECT GLC_AFFAIRELIEE FROM LIGNECOMPL ';
  Req := Req + 'WHERE GLC_NATUREPIECEG ="' + TOBPiece.getValue('GP_NATUREPIECEG') + '"';
  Req := Req + '  AND GLC_SOUCHE ="' + TOBPiece.getValue('GP_SOUCHE') + '"';
  Req := Req + '  AND GLC_NUMERO =' + intToStr(TOBPiece.getValue('GP_NUMERO'));
  Req := Req + '  AND GLC_INDICEG =' + intToStr(TOBPiece.getValue('GP_INDICEG'));

  TobGLC.LoadDetailDBFromSQL('LIGNECOMPL',req, false);

  if TOBGLC = nil then exit;

 	For I := 0 To TobGLC.Detail.count -1 Do
	    Begin
      naff := TOBGLC.Detail[i].GetValue('GLC_AFFAIRELIEE');
      if naff <> '' then
         Begin
         //Mise à jour du code état de l'appel
         Req := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="REA" ';
         Req := Req + 'WHERE AFF_AFFAIRE="' + Naff + '"';
         ExecuteSQL(Req);
         end;
      end;

  FreeAndNil(TOBGLC);

end;

procedure AFLanceFiche_Mul_Annu_PiecePro(Argument: string);
Var	Statut 	: string;
		Critere	: string;
    champ		: string;
    Valeur	: String;
    X				: Integer;
    SArg	  : String;
begin //mcd 30/07/03 passer de UtofAfpiece dans cette tof qui est concernée !!

//	  AGLLanceFiche('AFF', 'AFPIECEPROANU_MUL','AFF_AFFAIRE0=I;GP_VIVANTE=X', '', Argument);
	  AGLLanceFiche('AFF', 'AFPIECEPROANU_MUL','GP_VIVANTE=X', '', Argument);

end;


procedure AGLAnnulationFacture( Parms : array of variant ; nb : integer );
var  F : TForm;
     LaTof : TOF;
BEGIN

F:=TForm(Longint(Parms[0])) ;

if (F is TFmul) then
   LaTof:=TFMul(F).LaTOF
else
   exit;

if (LaTof is TOF_AfPieceProAnu_Mul) then
   TOF_AfPieceProAnu_Mul(LaTof).AnnulationFacture
else
	 exit;

END;

procedure InitAfPieceProAnu();
begin
RegisterAglProc('AnnulationFacture',True,1,AGLAnnulationFacture);
end;

{procedure TOF_AFPIECEPROANU_MUL.bSelectAff1Click(Sender: TObject);
begin

     if Statut = 'APP' then
        Editaff0.Text := 'W'
     Else if Statut = 'INT' then
        Editaff0.Text := 'I'
     Else if Statut = 'AFF' then
        Editaff0.Text := 'A'
     Else if Statut = 'PRO' then
        Editaff0.Text := 'P'
     else
        Editaff0.Text := '';

    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, VH_GC.GASeria , false, '', false, true, true);

end;}

procedure TOF_AFPIECEPROANU_MUL.DeflagueLesEcheanceslignes(TOBpiece: TOB);
var refpiece,Req : string;
begin
  Refpiece :=  EncodeRefPiece(TOBPIECE);
  Req := 'UPDATE FACTAFF SET AFA_ECHEFACT = "-",AFA_NUMPIECE =""'
    +',AFA_ETATVISA="",AFA_VISEUR="",AFA_DATEVISA="'+UsDateTime(idate1900)+'"'
    +' WHERE AFA_NUMPIECE='+'"'+RefPiece+'"';
  ExecuteSql (req);
end;

Initialization
registerclasses([TOF_AFPIECEPROANU_MUL]);
InitAfPieceProAnu();
// report modif PCS 13/05/03 Finalization
end.


