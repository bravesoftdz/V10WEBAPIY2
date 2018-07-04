{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 19/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : POINTAGECVJ ()
Mots clefs ... : TOF;POINTAGECVJ
*****************************************************************}
Unit POINTAGECVJ_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, FE_Main, UTob, HStatus, PrintDBG, Paramsoc;

procedure CC_LanceFichePointageCompteVersJournal;

Type
  TOF_POINTAGECVJ = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    ggrl_Pointage : THGrid;
    btn_Imprimer : TButton;
    gTobListeCpteP : Tob; // Liste des comptes pointables
    procedure ChercheComptePointable;
    function VerifieCptePUnique(pTobEcriture : Tob) : Integer;
    function IsPointable(pszCompte : String) : boolean;
    procedure btn_ImprimerClick(Sender: TObject);
  end ;

Implementation

procedure CC_LanceFichePointageCompteVersJournal;
begin
  if (GetParamSocSecur('SO_POINTAGEJAL',False) = False) then AGLLanceFiche('CP','POINTCVERSJ','','','')
                                             else PgiInfo('Vous êtes déjà en pointage sur journal.','Bascule du pointage sur compte vers journal');
end;

procedure TOF_POINTAGECVJ.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_POINTAGECVJ.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_POINTAGECVJ.OnUpdate ;
var
	Q,Q1 : TQuery;
  i,j,iPointage,iPosition : Integer;
  TobEcriture : Tob;
  szSQL,szWhere,szRefPointage1,szRefPointage2, szMessage : String;
  bOK, bMessage : Boolean;
  dTotal,dPiece,dOldPiece : Double;
  dtDatePointage : TDateTime;
  C : TControl;

	procedure Rapport(szDescription : String);
  begin
    ggrl_Pointage.Row := ggrl_Pointage.RowCount-1;
    with ggrl_Pointage do begin
      Cells[0,ggrl_Pointage.Row] := Q1.FindField('E_EXERCICE').AsString;
			Cells[1,ggrl_Pointage.Row] := Q1.FindField('E_JOURNAL').AsString;
			Cells[2,ggrl_Pointage.Row] := Q1.FindField('E_DATECOMPTABLE').AsString;
			Cells[3,ggrl_Pointage.Row] := Q1.FindField('E_NUMEROPIECE').AsString;
			Cells[4,ggrl_Pointage.Row] := Q1.FindField('E_NUMLIGNE').AsString;
      Cells[5,ggrl_Pointage.Row] := Q1.FindField('E_GENERAL').AsString;
      Cells[6,ggrl_Pointage.Row] := Q1.FindField('E_AUXILIAIRE').AsString;
      Cells[7,ggrl_Pointage.Row] := Q1.FindField('E_REFINTERNE').AsString;
      Cells[8,ggrl_Pointage.Row] := Q1.FindField('E_LIBELLE').AsString;
      Cells[9,ggrl_Pointage.Row] := Q1.FindField('E_DEBIT').AsString;
      Cells[10,ggrl_Pointage.Row] := Q1.FindField('E_CREDIT').AsString;
      Cells[11,ggrl_Pointage.Row] := Q1.FindField('E_REFPOINTAGE').AsString;
      Cells[12,ggrl_Pointage.Row] := Q1.FindField('E_DATEPOINTAGE').AsString;
      Cells[13,ggrl_Pointage.Row] := szDescription;
    end;
		ggrl_Pointage.RowCount := ggrl_Pointage.RowCount+1;
    bOk := False;
	end;
begin
  Inherited ;
  TobEcriture := nil;
  C := GetControl('BValider'); TButton(C).Enabled := False;
  btn_Imprimer.Visible := True;
  bMessage := True; {dPiece := 0;} dOldPiece := 0;

  // Toutes les lignes qui ont une référence de pointage
  Q := OpenSql('SELECT DISTINCT E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_QUALIFPIECE FROM ECRITURE WHERE E_REFPOINTAGE<>"" AND E_ECRANOUVEAU="N" AND E_QUALIFPIECE ="N" ORDER BY E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE',True);

  InitMove(Q.RecordCount,'');
	while not (Q.EOF) do begin
    // Ecriture entière correspondant à la ligne
    szSQL := 'SELECT E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_GENERAL, E_AUXILIAIRE, E_DEBIT, E_CREDIT, E_QUALIFPIECE, E_REFPOINTAGE, E_DATEPOINTAGE, E_NUMECHE, E_REFINTERNE, E_LIBELLE FROM ECRITURE';
    szWhere := ' WHERE E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" AND E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" AND E_DATECOMPTABLE="'+USDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString+' AND E_QUALIFPIECE="'+Q.FindField('E_QUALIFPIECE').AsString+'"';
    szSQL := szSQL+szWhere+' ORDER BY E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE';
    Q1 := OpenSql(szSQL,True);

    TobEcriture := Tob.Create('ECRITURE',Nil,-1);
    TobEcriture.LoadDetailDB('ECRITURE','','',Q1,False,False);
    bOk := True;
    dPiece := TobEcriture.Detail[0].GetValue('E_NUMEROPIECE');

    if (TobEcriture.Detail.Count = 2) then begin // 2 lignes dans l'écriture
      szRefPointage1 := TobEcriture.Detail[0].GetValue('E_REFPOINTAGE');
      szRefPointage2 := TobEcriture.Detail[1].GetValue('E_REFPOINTAGE');

      if (((szRefPointage1<>'') and (szRefPointage2<>'')) and (szRefPointage1<>szRefPointage2)) then begin
        iPointage := VerifieCptePUnique(TobEcriture);
        if (dPiece = dOldPiece) then Q1.Next else Q1.First;
        doldPiece := TobEcriture.Detail[0].GetValue('E_NUMEROPIECE');

        // Les deux comptes sont pointables : Inverse les références de pointage
        if (iPointage=-2) then begin
          dtDatePointage := TobEcriture.Detail[0].GetValue('E_DATEPOINTAGE');
          TobEcriture.Detail[0].PutValue('E_REFPOINTAGE',szRefPointage2);
          TobEcriture.Detail[0].PutValue('E_DATEPOINTAGE',TobEcriture.Detail[1].GetValue('E_DATEPOINTAGE'));
          TobEcriture.Detail[1].PutValue('E_REFPOINTAGE',szRefPointage1);
          TobEcriture.Detail[1].PutValue('E_DATEPOINTAGE',dtDatePointage);
          Q.Next; // Pour ne pas passer 2 fois sur la même écriture
          end
        else begin
          Rapport('La ligne d''écriture à mettre à jour contient déjà une référence de pointage');
          bOk := False;
          if bMessage then begin
            szMessage := '0;'+Ecran.Caption+';La ligne d''écriture à mettre à jour contient déjà une référence de pointage.#10#13Voulez-vous continuer ?;W;YNL;Y;N;;;';
            case HShowMessage(szMessage,'','') of
              mrNo  : begin Ferme(Q1); TobEcriture.Free; Ferme(Q); FiniMove; Exit; end;
              mrAll : bMessage := False;
            end;
          end;
        end;
        end
      else if ((szRefPointage1<>'') and (szRefPointage2<>'')) then bOk := False
      else if (szRefPointage1 <> '') then begin
        TobEcriture.Detail[1].PutValue('E_REFPOINTAGE',szRefPointage1);
        TobEcriture.Detail[1].PutValue('E_DATEPOINTAGE',TobEcriture.Detail[0].GetValue('E_DATEPOINTAGE'));
        end
      else begin
        TobEcriture.Detail[0].PutValue('E_REFPOINTAGE',szRefPointage2);
        TobEcriture.Detail[0].PutValue('E_DATEPOINTAGE',TobEcriture.Detail[1].GetValue('E_DATEPOINTAGE'));
      end;
      end
    // 3 lignes et plus dans l'écriture
    else begin
      // Vérifie qu'il n'y a qu'un seul compte pointable dans l'écriture : Si Oui : Pointe toute l'écriture
      iPointage := VerifieCptePUnique(TobEcriture);
      if (iPointage>0) then begin
        for i := 0 to TobEcriture.Detail.Count-1 do begin
          if (i<>iPointage) then begin
            TobEcriture.Detail[i].PutValue('E_REFPOINTAGE',TobEcriture.Detail[iPointage].GetValue('E_REFPOINTAGE'));
            TobEcriture.Detail[i].PutValue('E_DATEPOINTAGE',TobEcriture.Detail[iPointage].GetValue('E_DATEPOINTAGE'));
          end;
        end;
      end
      // Si plusieurs comptes pointables
      else begin
        // Parcours l'écriture et additionne (dans dTotal) debit et credit
        iPosition := 0; dTotal := 0;
        for i := 0 to TobEcriture.Detail.Count-1 do begin
          szRefPointage1 := TobEcriture.Detail[i].GetValue('E_REFPOINTAGE');
          dTotal := dTotal+TobEcriture.Detail[i].GetValue('E_DEBIT')-TobEcriture.Detail[i].GetValue('E_CREDIT');
          // Dés qu'on tombe sur un compte pointable, 2 cas possible :
          if (IsPointable(TobEcriture.Detail[i].GetValue('E_GENERAL'))) then begin
            // On à une référence de pointage, 2 cas possible :
            if (szRefPointage1<>'') then begin
              // dTotal =  0 : Pointe toute les lignes précédentes
              if (dTotal = 0) then begin
                dtDatePointage := TobEcriture.Detail[i].GetValue('E_DATEPOINTAGE');
                for j := iPosition to i-1 do begin
                  TobEcriture.Detail[j].PutValue('E_REFPOINTAGE',szRefPointage1);
                  TobEcriture.Detail[j].PutValue('E_DATEPOINTAGE',dtDatePointage);
                end;
                iPosition := i+1;
              end
              // dTotal <> 0 : Affiche une erreur : Propose un pointage manuel ?
              else begin
                if (dPiece <> dOldPiece) then Rapport('Impossible de transférer le pointage');
                doldPiece := TobEcriture.Detail[i].GetValue('E_NUMEROPIECE');
              end;
            end
            // On a pas de référence de pointage : Réinitialise dTotal.
            else begin
              dTotal := 0;
            end;
          end;
        end;
      end;
    end;
    if bok then if (TobEcriture.UpdateDB(False) = False) then Rapport('Impossible de mettre à jour l''enregistrement');
    Ferme(Q1);
    MoveCur(False);
    Q.Next;
  end;
  TobEcriture.Free;
  Ferme(Q);
  FiniMove;
  SetParamsoc('SO_POINTAGEJAL',True);
end ;

procedure TOF_POINTAGECVJ.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_POINTAGECVJ.OnArgument (S : String ) ;
begin
  Inherited ;
  btn_Imprimer := TButton(GetControl('BIMPRIMER'));	if (btn_Imprimer <> nil) then btn_Imprimer.OnClick := btn_ImprimerClick;
  ggrl_Pointage := THGrid(GetControl('GRL_POINTAGE'));
  ggrl_Pointage.ColWidths[0] := 82;
  ggrl_Pointage.ColWidths[1] := 80;
  ggrl_Pointage.ColWidths[2] := 90;
  ggrl_Pointage.ColWidths[3] := 90;
  ggrl_Pointage.ColWidths[4] := 90;
  ggrl_Pointage.ColWidths[5] := 90;
  ggrl_Pointage.ColWidths[6] := 90;
  ggrl_Pointage.ColWidths[7] := 90;
  ggrl_Pointage.ColWidths[8] := 90;
  ggrl_Pointage.ColWidths[9] := 100;
  ggrl_Pointage.ColWidths[10] := 100;
  ggrl_Pointage.ColWidths[11] := 90;
  ggrl_Pointage.ColWidths[12] := 90;
  ggrl_Pointage.ColWidths[13] := 370;
  ggrl_Pointage.ColFormats[9] := '#,#00.00';
  ggrl_Pointage.ColFormats[10] := '#,#00.00';
  ChercheComptePointable;
end ;

procedure TOF_POINTAGECVJ.OnClose ;
begin
  Inherited ;

end ;

// Renseigne la liste des comptes pointables
procedure TOF_POINTAGECVJ.ChercheComptePointable;
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_POINTABLe="X"',True);
  gTobListeCpteP := Tob.Create('Vituelle',Nil,-1);
  gTobListeCpteP.LoadDetailDB('0','','',Q,False,False);
  Ferme(Q);
end;

// Vérifie s'il n'y a qu'un seul compte pointable dans l'écriture
function TOF_POINTAGECVJ.VerifieCptePUnique(pTobEcriture: Tob): Integer;
var
  i,j,iCompteur : Integer;
  szCompte : String;
begin
  Result := -1; iCompteur := 0;
  for i := 0 to pTobEcriture.Detail.Count-1 do begin
    szCompte := pTobEcriture.Detail[i].GetValue('E_GENERAL');
    for j := 0 to gTobListeCpteP.Detail.Count -1 do begin
      if szCompte = gTobListeCpteP.Detail[j].GetValue('G_GENERAL') then begin
        inc(iCompteur);
        if (iCompteur = 1) then Result := i;
        break;
      end;
    end;
    if (iCompteur = 2) then begin Result := -2; Exit; end;
  end;
end;

// Vérifie si le compte est pointable
function TOF_POINTAGECVJ.IsPointable(pszCompte: String): boolean;
var
  i : Integer;
begin
  Result := False;
  for i := 0 to gTobListeCpteP.Detail.Count -1 do
    if pszCompte = gTobListeCpteP.Detail[i].GetValue('G_GENERAL') then begin Result := True; Exit; end;
end;

procedure TOF_POINTAGECVJ.btn_ImprimerClick(Sender: TObject);
begin
  PrintDBGrid(ggrl_Pointage,Nil,Ecran.Caption,'');
end;

Initialization
  registerclasses ( [ TOF_POINTAGECVJ ] ) ;
end.
