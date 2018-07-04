{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/06/2002
Modifié le ... : 20/06/2002
Description .. : Source TOF de la FICHE : ECARTCONV ()
Mots clefs ... : TOF;ECARTCONV
*****************************************************************}
Unit ECARTCONV_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db, dbtables,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, FE_Main,utilPGI, Ent1, UTob, HStatus, PrintDBG, HPanel, ParamSoc;

procedure CCLanceFiche_SupprEcartConv;

Type
  TOF_ECARTCONV = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
	private
  	btn_Imprimer : TButton;
    grl_Rapport : THGrid;
    procedure btn_ImprimerClick(Sender: TObject);
    function ExoEstOuvert(szExo : String) : boolean;
  end ;

Implementation

const DateFinFranc : String = '01/01/2002';

{=======================================================================================}
procedure CCLanceFiche_SupprEcartConv;
begin
	if Not _BlocageMonoPoste(True,'',TRUE) then Exit ;
	AGLLanceFiche('CP','CPECARTCONV','','','');
end ;

{=======================================================================================}
procedure TOF_ECARTCONV.OnUpdate ;
var
	Q,Q1,Q2 : TQuery;
  TobEcriture : Tob;
  TobAnalytique : Tob;
  iRetour,i,j,iNumLigne,iNumLigneAna,iCompteur : Integer;
  S,stGeneral,stAdd1,stAdd2,szSet1,szSet2 : String;
  dDebit, dCredit : double;

	procedure Rapport(szDescription : String);
  begin
    grl_Rapport.Row := grl_Rapport.RowCount-1;
    with grl_Rapport do begin
			Cells[0,grl_Rapport.Row] := Q.FindField('E_JOURNAL').AsString;
			Cells[1,grl_Rapport.Row] := Q.FindField('E_DATECOMPTABLE').AsString;
			Cells[2,grl_Rapport.Row] := Q.FindField('E_NUMEROPIECE').AsString;
			Cells[3,grl_Rapport.Row] := Q.FindField('E_NUMLIGNE').AsString;
			Cells[4,grl_Rapport.Row] := Q.FindField('E_GENERAL').AsString;
			Cells[5,grl_Rapport.Row] := Q.FindField('E_AUXILIAIRE').AsString;
			Cells[6,grl_Rapport.Row] := Q.FindField('E_REFINTERNE').AsString;
      Cells[7,grl_Rapport.Row] := Q.FindField('E_LIBELLE').AsString;
			Cells[8,grl_Rapport.Row] := Q.FindField('E_DEBIT').AsString;
			Cells[9,grl_Rapport.Row] := Q.FindField('E_CREDIT').AsString;
			Cells[10,grl_Rapport.Row] := Q.FindField('E_DEBITEURO').AsString;
			Cells[11,grl_Rapport.Row] := Q.FindField('E_CREDITEURO').AsString;
      Cells[12,grl_Rapport.Row] := szDescription;
    end;
		grl_Rapport.RowCount := grl_Rapport.RowCount+1;
	end;
begin
  Inherited ;
	// Ne peut être lancé que si bascule euro est faite
  if not VH^.TenueEuro then exit;

  // Compte d'écart de conversion ventilables
  if ExisteSQL('SELECT * FROM GENERAUX WHERE (G_GENERAL="'+VH^.ECcEuroDebit+'" OR G_GENERAL="'+VH^.ECcEuroCredit+'") AND G_VENTILABLE="X"') then begin
    HShowMessage('0;'+Ecran.caption+';Impossible de lancer le traitement. Les comptes d''écart de conversion sont ventilables.;W;O;O','','');
    Exit;
  end;

  // Compte d'écart de conversion identiques aux comptes d'écart de change
  if ExisteSQL('SELECT * FROM DEVISE WHERE D_CPTLETTRDEBIT="'+VH^.ECcEuroDebit+'" OR  D_CPTLETTRCREDIT="'+VH^.ECcEuroDebit+'" OR D_CPTPROVDEBIT="'+VH^.ECcEuroDebit+'" OR D_CPTPROVCREDIT="'+VH^.ECcEuroDebit+'" OR D_CPTLETTRDEBIT="'+VH^.ECcEuroCredit+'" OR  D_CPTLETTRCREDIT="'+VH^.ECcEuroCredit+'" OR D_CPTPROVDEBIT="'+VH^.ECcEuroCredit+'" OR D_CPTPROVCREDIT="'+VH^.ECcEuroCredit+'"') then begin
    HShowMessage('0;'+Ecran.caption+';Impossible de lancer le traitement. Une ou plusieurs devises utilisent les comptes d''écart de conversion.;W;O;O','','');
    Exit;
  end;

  if (not TCheckBox(GetControl('CHK_JALEXTRACPT')).Checked) then stAdd1 := ' AND J_NATUREJAL<>"EXT"';
  if (not TCheckBox(GetControl('CHK_JALREGUL')).Checked) then stAdd2 := ' AND J_NATUREJAL<>"REG"';

  // Ecarts de conversion générés lors de la saisie en mode pièce (E_MODESAISIE = '-' ou '')
  // La pièce est équilibré au jour. Toutes les lignes de la pièce ont le même n° de pièce, datecomptable, journal, etablissement, devise
	if (TCheckBox(GetControl('CHK_MODEPIECE')).Checked) then begin
		Q := OpenSql('SELECT E_EXERCICE,E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_DEBIT,E_CREDIT,E_QUALIFPIECE,E_NUMECHE FROM ECRITURE LEFT JOIN JOURNAL ON E_JOURNAL = J_JOURNAL WHERE E_DATECOMPTABLE >="'+DateFinFranc+'" AND (E_MODESAISIE="-" OR E_MODESAISIE="") AND (E_GENERAL="'+VH.EccEuroDebit+'" OR E_GENERAL="'+VH.EccEuroCredit+'") AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ECC"'+stAdd1+stAdd2,True);
    InitMove(Q.RecordCount,'');
		while not (Q.EOF) do begin
    	if (Q.FindField('E_DEBIT').AsFloat = Q.FindField('E_CREDIT').AsFloat) then
        if ExoEstOuvert(Q.FindField('E_EXERCICE').AsString) then
      		if (ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE ="'+Q.FindField('E_EXERCICE').AsString+'" AND E_JOURNAL ="'+Q.FindField('E_JOURNAL').AsString+'" AND E_DATECOMPTABLE ="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND E_NUMEROPIECE ='+Q.FindField('E_NUMEROPIECE').AsString+' AND E_NUMLIGNE ='+Q.FindField('E_NUMLIGNE').AsString+' AND E_QUALIFPIECE ="'+Q.FindField('E_QUALIFPIECE').AsString+'" AND E_NUMECHE ='+Q.FindField('E_NUMECHE').AsString) = 0) then Rapport('Impossible de supprimer l''enregistrement');
        MoveCur(False);
      Q.Next;
    end;
    Ferme(Q);
    TCheckBox(GetControl('CHK_MODEPIECE')).Checked := False;
    FiniMove;
  end;

  // Ecarts de conversion générés lors de la saisie en mode borderau (E_MODESAISIE = 'BOR')
  // La pièce est équilibré au jour. Touts les lignes de la pièce ont le même n° de pièce, période, journal, etablissement, devise
  // Après suppression des écritures : Renuméroter toutes les autres sans oublier l'analytique
	if (TCheckBox(GetControl('CHK_MODEBORDEREAU1')).Checked) then begin
	  Q := OpenSql('SELECT DISTINCT E_NUMEROPIECE, E_JOURNAL FROM ECRITURE LEFT JOIN JOURNAL ON E_JOURNAL = J_JOURNAL WHERE E_DATECOMPTABLE >="'+DateFinFranc+'" AND E_MODESAISIE="BOR" AND (E_GENERAL="'+VH.EccEuroDebit+'" OR E_GENERAL="'+VH.EccEuroCredit+'") AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ECC"'+stAdd1+stAdd2,True);
    while not (Q.EOF) do begin
			// Tob les pièces de la Table Ecriture
      Q1 := OpenSql('SELECT E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_QUALIFPIECE, E_NUMECHE, E_GENERAL FROM ECRITURE WHERE E_DATECOMPTABLE >="'+DateFinFranc+'" AND E_MODESAISIE="BOR" AND E_NUMEROPIECE="'+Q.FindField('E_NUMEROPIECE').AsString+'" AND E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"',True);
      TobEcriture := Tob.Create('ECRITURE',Nil,-1);
      TobEcriture.LoadDetailDB('ECRITURE','','',Q1,False,False);

      // Tob les pièces de la Table Analytiq
      Q2 := OpenSql('SELECT Y_AXE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_DEBIT, Y_CREDIT, Y_EXERCICE, Y_QUALIFPIECE, Y_JOURNAL, Y_NUMVENTIL FROM ANALYTIQ WHERE Y_DATECOMPTABLE >="'+DateFinFranc+'" AND Y_NUMEROPIECE="'+Q.FindField('E_NUMEROPIECE').AsString+'" AND Y_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"',True);
      TobAnalytique := Tob.Create('ANALYTIQ',Nil,-1);
      TobAnalytique.LoadDetailDB('ANALYTIQ','','',Q2,False,False);

      // Parcours la tob ecriture à la recherche d'un compte général = VH^.ECcEuroDebit ou VH^.ECcEuroCredit
      InitMove(TobEcriture.Detail.Count,'');
      i := 0; iCompteur := 0;
      while (i < TobEcriture.Detail.Count) do begin
	      stGeneral := TobEcriture.Detail[i].GetValue('E_GENERAL');
        iNumLigne := TobEcriture.Detail[i].GetValue('E_NUMLIGNE');
        dDebit := TobEcriture.Detail[i].GetValue('E_DEBIT');
        dCredit := TobEcriture.Detail[i].GetValue('E_CREDIT');
      	if ((stGeneral = VH^.ECcEuroDebit) or (stGeneral = VH^.ECcEuroCredit)) and (dDebit = dCredit) then begin	// Ligne d'écart de conversion
          if ExoEstOuvert(TobEcriture.Detail[i].GetValue('E_EXERCICE')) then begin
	          inc(iCompteur);

            // Rénumérote les lignes analytique
            for j := 0 to TobAnalytique.Detail.Count-1 do begin
	            iNumLigneAna := TobAnalytique.Detail[j].GetValue('Y_NUMLIGNE');
            	if (iNumLigneAna > iNumLigne) then begin
					  		TobAnalytique.Detail[j].PutValue('Y_NUMLIGNE',iNumLigneAna - 1);
              end
          	end;
            TobEcriture.Detail[i].DeleteDB; // Supprime l'enregistrement
            TobEcriture.Detail[i].Free;
            dec(i);
          end;
        end;
        if (iCompteur>0) then TobEcriture.Detail[i].PutValue('E_NUMLIGNE',iNumLigne-iCompteur);
        inc(i);
      end;

     	// Met à jour les tob
      TobEcriture.UpdateDB(False);
      TobAnalytique.UpdateDB(False);
      Ferme(Q1); Ferme(Q2);
      Q.Next;
      MoveCur(False);
      TobEcriture.Free;
	    TobAnalytique.Free;
    end;
		Ferme(Q);
    TCheckBox(GetControl('CHK_MODEBORDEREAU1')).Checked := False;
    FiniMove;
  end;

  // Ecarts de conversion générés lors de la saisie en mode bordereau (E_MODESAISIE = 'LIB')
  // La pièce est équilibré au mois. Toutes les lignes de la pièce ont le même n° de pièce, période, journal, etablissement, devise
  // Après suppression des écritures : Renuméroter toutes les autres sans oublier l'analytique
  if (TCheckBox(GetControl('CHK_MODEBORDEREAU2')).Checked) then begin
	  Q := OpenSql('SELECT DISTINCT E_NUMEROPIECE, E_JOURNAL FROM ECRITURE LEFT JOIN JOURNAL ON E_JOURNAL = J_JOURNAL WHERE E_DATECOMPTABLE >="'+DateFinFranc+'" AND E_MODESAISIE="LIB" AND (E_GENERAL="'+VH.EccEuroDebit+'" OR E_GENERAL="'+VH.EccEuroCredit+'") AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ECC"'+stAdd1+stAdd2,True);
    while not (Q.EOF) do begin
			// Tob les pièces de la Table Ecriture
      Q1 := OpenSql('SELECT E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_QUALIFPIECE, E_NUMECHE, E_GENERAL FROM ECRITURE WHERE E_DATECOMPTABLE >="'+DateFinFranc+'" AND E_MODESAISIE="LIB" AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString+' AND E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" ORDER BY E_NUMLIGNE',True);
      TobEcriture := Tob.Create('ECRITURE',Nil,-1);
      TobEcriture.LoadDetailDB('ECRITURE','','',Q1,False,False);

      // Tob les pièces de la Table Analytiq
      Q2 := OpenSql('SELECT Y_AXE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_EXERCICE, Y_QUALIFPIECE, Y_JOURNAL, Y_NUMVENTIL FROM ANALYTIQ WHERE Y_DATECOMPTABLE >="'+DateFinFranc+'" AND Y_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString+' AND Y_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" ORDER BY Y_NUMLIGNE',True);
      TobAnalytique := Tob.Create('ANALYTIQ',Nil,-1);
      TobAnalytique.LoadDetailDB('ANALYTIQ','','',Q2,False,False);

      // Parcours la tob ecriture à la recherche d'un compte général = VH^.ECcEuroDebit ou VH^.ECcEuroCredit
      InitMove(TobEcriture.Detail.Count,'');
      i := 0; iCompteur := 0;
      while (i < TobEcriture.Detail.Count) do begin
	      stGeneral := TobEcriture.Detail[i].GetValue('E_GENERAL');
        iNumLigne := TobEcriture.Detail[i].GetValue('E_NUMLIGNE');
        dDebit := TobEcriture.Detail[i].GetValue('E_DEBIT');
        dCredit := TobEcriture.Detail[i].GetValue('E_CREDIT');
      	if ((stGeneral = VH^.ECcEuroDebit) or (stGeneral = VH^.ECcEuroCredit)) and (dDebit = dCredit) then begin	// Ligne d'écart de conversion
          if ExoEstOuvert(TobEcriture.Detail[i].GetValue('E_EXERCICE')) then begin
            inc(iCompteur);

            // Rénumérote les lignes analytique
            for j := 0 to TobAnalytique.Detail.Count-1 do begin
	            iNumLigneAna := TobAnalytique.Detail[j].GetValue('Y_NUMLIGNE');
            	if (iNumLigneAna > iNumLigne) then begin
						  	TobAnalytique.Detail[j].PutValue('Y_NUMLIGNE',iNumLigneAna - 1);
              end
          	end;
            TobEcriture.Detail[i].DeleteDB; // Supprime l'enregistrement
            TobEcriture.Detail[i].Free;
            dec(i);
          end;
        end;
        if (iCompteur>0) then TobEcriture.Detail[i].PutValue('E_NUMLIGNE',iNumLigne-iCompteur);
        inc(i);
      end;

     	// Met à jour les tob
      TobEcriture.UpdateDB(False);
      TobAnalytique.UpdateDB(False);
      Ferme(Q1); Ferme(Q2);
      Q.Next;
      MoveCur(False);
      TobEcriture.Free;
	    TobAnalytique.Free;
    end;
		Ferme(Q);
    TCheckBox(GetControl('CHK_MODEBORDEREAU2')).Checked := False;
    FiniMove;
  end;

  // Ecarts de conversion générés lors du lettrage
  if (TCheckBox(GetControl('CHK_PIECES')).Checked) then begin
  	Q := OpenSql('SELECT E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_GENERAL, E_AUXILIAIRE, E_REFINTERNE, E_LIBELLE, E_QUALIFPIECE, E_NUMECHE, E_LETTRAGE, E_DEBIT, E_CREDIT, E_DEBITEURO, E_CREDITEURO FROM ECRITURE'+' WHERE E_ETATLETTRAGE ="TL" AND E_DATECOMPTABLE>="'+DateFinFranc+'" AND E_JOURNAL="'+VH^.JalEcartEuro+'"',True);
    InitMove(Q.RecordCount,'');
    while not (Q.EOF) do begin
			// Tob les pièces de la Table Ecriture
//      Q1 := OpenSql('SELECT SUM(E_DEBIT) AS E_DEBIT, SUM(E_CREDIT) AS E_CREDIT FROM ECRITURE WHERE E_ETATLETTRAGE ="TL" AND E_LETTRAGE="'+Q.FindField('E_LETTRAGE').AsString+'" AND E_JOURNAL<>"'+VH^.JalEcartEuro+'" AND E_AUXILIAIRE="'+Q.FindField('E_AUXILIAIRE').AsString+'"',True);
      Q1 := OpenSql('SELECT E_EXERCICE,E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_DEBIT,E_CREDIT,E_QUALIFPIECE,E_NUMECHE FROM ECRITURE WHERE E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" AND E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" AND E_DATECOMPTABLE="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString,True);
      TobEcriture := Tob.Create('ECRITURE',Nil,-1);
      TobEcriture.LoadDetailDB('ECRITURE','','',Q1,False,False);

      // Additionne le débit et crédit
      dDebit := 0; dCredit := 0;
      for i := 0 to TobEcriture.Detail.Count-1 do begin
        dDebit := dDebit+TobEcriture.Detail[i].GetValue('E_DEBIT');
        dCredit := dCredit+TobEcriture.Detail[i].GetValue('E_CREDIT');
      end;

      // Si Total Debit = Total Credit : Supprime l'écriture d'écart de conversion uniquement si l'exercice n'est pas clôturé
      if (Abs(dDebit-dCredit)<0.001) then begin
        if ExoEstOuvert(TobEcriture.Detail[0].GetValue('E_EXERCICE')) then begin
//				Q2 := OpenSql('SELECT EX_ETATCPTA FROM EXERCICE WHERE EX_EXERCICE ="'+Q.FindField('E_EXERCICE').AsString+'"',True);
//        if (Q2.FindField('EX_ETATCPTA').AsString = 'OUV') then begin
//					if (ExecuteSQL('DELETE FROM ECRITURE WHERE E_EXERCICE ="'+Q.FindField('E_EXERCICE').AsString+'" AND E_JOURNAL ="'+Q.FindField('E_JOURNAL').AsString+'" AND E_DATECOMPTABLE ="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND E_NUMEROPIECE ="'+Q.FindField('E_NUMEROPIECE').AsString+'" AND E_NUMLIGNE ="'+Q.FindField('E_NUMLIGNE').AsString+'" AND E_QUALIFPIECE ="'+Q.FindField('E_QUALIFPIECE').AsString+'" AND E_NUMECHE ="'+Q.FindField('E_NUMECHE').AsString+'"') = 0) then Rapport('Impossible de supprimer l''enregistrement');
          szSet1 := 'E_CREDIT=0, E_REFINTERNE="REMISE A 0 DES ECARTS DE CONVERSION", E_LIBELLE="DETRUITE LE '+FormatDateTime('dd/mm/yyyy',Date)+'", E_VALIDE="X", E_REFEXTERNE="", E_DATEREFEXTERNE="01/01/1900", E_BUDGET="", E_CONTREPARTIEGEN="", E_CONTREPARTIEAUX="", ';
          szSet1 := szSet1+'E_COUVERTURE=0, E_LETTRAGE="", E_LETTRAGEDEV="-", E_REFPOINTAGE="", E_DATEPOINTAGE="01/01/1900", E_DATEECHEANCE="01/01/1900", E_DATERELANCE="01/01/1900", E_NIVEAURELANCE=0, ';
          szSet2 := 'E_CREDITEURO=0, E_DEBITDEV=0, E_CREDITDEV=0, E_TAUXDEV=1, E_RIB="", E_REFRELEVE="", E_COUVERTUREDEV=0, E_ETATLETTRAGE="AL", E_NUMPIECEINTERNE="", E_NOMLOT="", E_TRACE="", E_NUMECHE=0, E_CREERPAR="DET", E_COUVERTUREEURO=0, E_SAISIEEURO="X", ';
          szSet2 := szSet2+'E_LETTRAGEEURO="X", E_LIBRETEXTE0="", E_LIBRETEXTE1="",  E_LIBRETEXTE2="", E_LIBRETEXTE3="", E_LIBRETEXTE4="", E_LIBRETEXTE5="", E_LIBRETEXTE6="", E_LIBRETEXTE7="", E_LIBRETEXTE8="", E_LIBRETEXTE9="", E_TABLE0="", E_TABLE1="", ';
          szSet2 := szSet2+'E_LIBREDATE="01/01/1900", E_LIBREBOOL0="-", E_LIBREBOOL1="-", E_CONSO="", E_LIBREMONTANT0=0, E_LIBREMONTANT1=0, E_LIBREMONTANT2=0, E_LIBREMONTANT3=0, E_REFLETTRAGE="", E_TABLE2="", E_TABLE3="", E_PIECETP=""';
          S:= 'E_EXERCICE ="'+Q.FindField('E_EXERCICE').AsString+'" AND E_JOURNAL ="'+Q.FindField('E_JOURNAL').AsString+'" AND E_DATECOMPTABLE ="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND E_NUMEROPIECE ='+Q.FindField('E_NUMEROPIECE').AsString+' AND E_QUALIFPIECE ="'+Q.FindField('E_QUALIFPIECE').AsString+'"';
          BeginTrans ;
          iRetour := ExecuteSQL('DELETE FROM ECRITURE WHERE '+S+'AND (E_NUMLIGNE>2 OR E_NUMECHE>1)');
          iRetour := iRetour + ExecuteSQL('UPDATE ECRITURE SET E_GENERAL="'+GetParamSocSecur('SO_GENATTEND','')+'",E_AUXILIAIRE="", E_DEBIT=1, '+szSet1+'E_DEBITEURO=1, '+SzSet2+' WHERE '+S+' AND E_NUMLIGNE=1');
          iRetour := iRetour + ExecuteSQL('UPDATE ECRITURE SET E_GENERAL="'+GetParamSocSecur('SO_GENATTEND','')+'",E_AUXILIAIRE="", E_DEBIT=-1, '+szSet1+'E_DEBITEURO=-1, '+SzSet2+' WHERE '+S+' AND E_NUMLIGNE=2');
          if (iRetour = 0) then begin Rapport('Impossible de détruire l''enregistrement'); Rollback; end
                           else CommitTrans;
          end
        else Rapport('Exercice clos');
//        Ferme(Q2);
      	end
      else Rapport('Déséquilibre de l''écriture');

     	// Met à jour les tob
      Ferme(Q1);
      Q.Next;
      MoveCur(False);
      TobEcriture.Free;
    end;
    Ferme(Q);
    TCheckBox(GetControl('CHK_PIECES')).Checked := False;
    FiniMove;
  end;

	// Rapport sur les erreurs
	if grl_Rapport.RowCount > 2 then begin
		btn_Imprimer.Visible := True;
    Ecran.caption := Ecran.caption + ' : Détail des écritures non supprimées'; UpdateCaption(Ecran);
    THPanel(GetControl('PNL_RAPPORT')).Visible := True;
  end;
end ;

{=======================================================================================}
procedure TOF_ECARTCONV.OnLoad ;
begin
  Inherited ;
end ;

{=======================================================================================}
procedure TOF_ECARTCONV.OnArgument (S : String ) ;
begin
  Inherited ;
  btn_Imprimer := TButton(GetControl('BIMPRIMER'));	if (btn_Imprimer <> nil) then btn_Imprimer.OnClick := btn_ImprimerClick;
  grl_Rapport := THGrid(GetControl('GRL_RAPPORT'));
end ;

{=======================================================================================}
procedure TOF_ECARTCONV.OnClose ;
begin
  Inherited ;
  _DeblocageMonoPoste(True,'',TRUE) ;
end ;

{=======================================================================================}
procedure TOF_ECARTCONV.btn_ImprimerClick(Sender: TObject);
begin
	if (THPanel(GetControl('PNL_RAPPORT')).Visible = True) then
		PrintDBGrid(grl_Rapport,Nil,Ecran.Caption,'');
end;

{=======================================================================================}
function TOF_ECARTCONV.ExoEstOuvert(szExo: String): boolean;
begin
  Result := ExisteSQL('SELECT EX_EXERCICE FROM EXERCICE WHERE EX_EXERCICE="'+szExo+'" AND EX_ETATCPTA="OUV"');
end;

{=======================================================================================}
Initialization
  registerclasses ( [ TOF_ECARTCONV ] ) ;
end.
