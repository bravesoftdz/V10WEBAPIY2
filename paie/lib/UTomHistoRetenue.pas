{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 23/06/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : HISTORETENUE (HISTORETENUE)
Mots clefs ... : TOM;HISTORETENUE
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
}
Unit UTomHistoRetenue;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,EdtREtat,
{$ELSE}
     eFiche,UtileAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,SaisieList,uTableFiltre,HTB97 ;

Type
  TOM_HISTORETENUE = Class (TOM)
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
    Salarie,NumOrdre : String;
    TF: TTableFiltre;
    MtTotalRetenue : Double;
    procedure BEtatClick (Sender : TObject);
    end ;

Implementation

procedure TOM_HISTORETENUE.OnUpdateRecord ;
var DateDebut,DateFin : TDateTime;
    MtVerse : Double;
    Q : TQuery;
begin
  Inherited ;
{        DateDebut := GetField('PHR_DATEDEBUT');
        DateFin := GetField('PHR_DATEFIN');
        If TF.State = DsInsert then
        begin
                If ExisteSQL('SELECT PHR_SALARIE FROM HISTORETENUE WHERE PHR_SALARIE="'+Salarie+'" AND PHR_ORDRE='+NumOrdre+''+
                ' AND ((PHR_DATEDEBUT<="'+UsDateTime(DateDebut)+'" AND PHR_DATEFIN>="'+UsDateTime(DateDebut)+'") OR '+
                '(PHR_DATEDEBUT<="'+UsDateTime(DateFin)+'" AND PHR_DATEFIN>="'+UsDateTime(DateFin)+'"))') then
                begin
                        LastError := 1;
                        PGIBox('Saisie impossible, cat la période choisie est inscluse dans une autre période', Ecran.Caption);
                end;
        end;}
        If GetField('PHR_EFFCTUE') = '-' then SetField('PHR_DATEPAIEMENT',IDate1900);
        If (GetField('PHR_EFFECTUE') = 'X') and (GetField('PHR_DATEPAIEMENT') <= IDate1900) then
        begin
                LastError := 1;
                PGIbox('Vous devez saisir une date de paiement',Ecran.Caption);
                Exit;
        end;
        If GetField('PHR_EFFECTUE') = 'X' then
        begin
                Q := OpenSQL('SELECT SUM(PHR_MONTANT) CUMULVERSE FROM HISTORETENUE '+
                      'WHERE PHR_SALARIE="' + GetField('PHR_SALARIE') +
                      '" AND PHR_ORDRE=' + IntToStr(GetField('PHR_ORDRE')) + 'AND '+
                      '(PHR_DATEDEBUT <>"' + UsDateTime(GetField('PHR_DATEDEBUT')) + '" '+
                      'OR PHR_DATEPAIEMENT <>"' + UsDateTime(GetField('PHR_DATEFIN')) + '") AND PHR_EFFECTUE="X"', True);
                If Not Q.Eof then MtVerse := Q.FindField('CUMULVERSE').AsFloat
                else MtVerse := 0;
                Ferme(Q);
                MtVerse := MtVerse + GetField('PHR_MONTANT');
                If MtVerse > MtTotalRetenue then
                begin
                        LastError := 1;
                        PGIbox('Le montant des paiements ('+FloatToStr(MtVerse)+') ne peut être supérieur '+
                        'au montant total de la retenue ('+FloatToStr(MtTotalRetenue),Ecran.Caption);
                        Exit;
                end;
        end;
     If not (Ecran is TFSaisieList) Then
     begin
      SetField('PHR_DATEDEBUT',GetField('PHR_DATEPAIEMENT'));
      SetField('PHR_DATEFIN',GetField('PHR_DATEPAIEMENT'));
     end;
end;

procedure TOM_HISTORETENUE.OnAfterUpdateRecord;
var DatePaiement,rubrique : String;
    TobMajHistorique,Tmaj : Tob;
    CumulVerse,Paye,CumulArriere,Arriere,MtMens,CumulMens,Reprise : Double;
    i : Integer;
begin
  Inherited ;
  If not (Ecran is TFSaisieList) Then Exit;
       // if GetField('PHR_EFFECTUE') = 'X' then
       // begin
                If GetField('PHR_RETENUESAL') = '001' then Rubrique := '7230';
                If GetField('PHR_RETENUESAL') = '002' then Rubrique := '7240';
                If GetField('PHR_RETENUESAL') = '003' then Rubrique := '7250';
                If GetField('PHR_RETENUESAL') = '004' then Rubrique := '7260';
                If Not ExisteSQL('SELECT PHR_SALARIE FROM HISTORETENUE '+
                'LEFT JOIN HISTOBULLETIN ON PHB_DATEDEBUT=PHR_DATEDEBUT AND PHB_DATEFIN=PHR_DATEFIN '+
                'AND PHB_SALARIE=PHR_SALARIE WHERE PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+Rubrique+'" '+
                'AND PHB_SALARIE="'+GetField('PHR_SALARIE')+'" AND PHB_DATEDEBUT="'+UsDateTime(GetField('PHR_DATEDEBUT'))+'" '+
                'AND PHB_DATEFIN="'+UsDateTime(GetField('PHR_DATEFIN'))+'" AND PHB_MTREM='+IntToStr(GetField('PHR_MONTANT'))+'') then
                begin
                        TobMajHistorique := Tob.Create('MajHisto',Nil,-1);
                        DatePaiement := GetField('PHR_DATEPAIEMENT');
                        TF.StartUpdate;
                        TF.DisableTOM;
                        for i:=1 to TF.LaGrid.RowCount - 1 do
                        begin
                                TF.SelectRecord(i);
                                TMaj := Tob.Create('FilleMaj',TobMajHistorique,-1);
                                TMaj.AddChampSupValeur('NUMORDRE',TF.GetValue('PHR_ORDRE'),False);
                                TMaj.AddChampSupValeur('DDEB',TF.GetValue('PHR_DATEDEBUT'),False);
                                TMaj.AddChampSupValeur('DFIN',TF.GetValue('PHR_DATEFIN'),False);
                                TMaj.AddChampSupValeur('DPAIE',TF.GetValue('PHR_DATEPAIEMENT'));
                                TMaj.AddChampSupValeur('MTMENS',TF.GetValue('PHR_MONTANTMENS'));
                                Paye := TF.GetValue('PHR_MONTANT');
                                TMaj.AddChampSupValeur('MTPAYE',Paye);
                                TMaj.AddChampSupValeur('ARRIERE',0);
                                TMaj.AddChampSupValeur('PHR_REPRISEARR',0);
                                TMaj.AddChampSupValeur('CUMULARRIERE',0);
                                TMaj.AddChampSupValeur('CUMULVERSE',0);
                                TMaj.AddChampSupValeur('EFFECTUE',TF.GetValue('PHR_EFFECTUE'));
                        end;
                        TobMajHistorique.Detail.Sort('DPAIE');
                        CumulVerse := 0;
                        CumulArriere := 0;
                        CumulMens := 0;
                        For i := 0 to TobMajHistorique.Detail.Count - 1 do
                        begin
                                Tmaj := TobMajHistorique.Detail[i];
                               If TMaj.GetValue('EFFECTUE') = 'X' then
                               begin
                                       Paye := TMaj.GetValue('MTPAYE');
                                       MtMens := TMaj.GetValue('MTMENS');
                                       CumulMens := CumulMens + MtMens;
                                       CumulVerse := CumulVerse + Paye;
                                       Arriere := MtMens - Paye;
                                       If Arriere < 0 then
                                       begin
                                        Reprise := - Arriere;
                                        Arriere := 0;
                                       end;
                                       CumulArriere := CumulArriere + Arriere;
                                       TMaj.PutValue('CUMULVERSE',CumulVerse);
                                       TMaj.PutValue('ARRIERE',Arriere);
                                       TMaj.PutValue('CUMULARRIERE',CumulArriere);
                                       TMaj.PutValue('PHR_REPRISEARR',Reprise);
                               end
                               else
                               begin
                                       TMaj.PutValue('CUMULVERSE',0);
                                       TMaj.PutValue('PHR_ARRIERE',0);
                                       TMaj.PutValue('PHR_CUMULARRIERE',0);
                                       TMaj.PutValue('PHR_REPRISEARR',0);
                               end;
                        end;
                        for i:=1 to TF.LaGrid.RowCount - 1 do
                        begin
                                TF.SelectRecord(i);
                                TMaj := TobMajHistorique.findFirst(['DDEB','DFIN'],[TF.GetValue('PHR_DATEDEBUT'),TF.GetValue('PHR_DATEFIN')],False);
                                If TMaj <> Nil then
                                begin
                                        TF.PutValue('PHR_CUMULVERSE',TMaj.GetValue('CUMULVERSE'));
                                        TF.PutValue('PHR_ARRIERE',TMaj.GetValue('ARRIERE'));
                                        TF.PutValue('PHR_CUMULARRIERE',TMaj.GetValue('CUMULARRIERE'));
                                        TF.PutValue('PHR_REPRISEARR',TMaj.GetValue('PHR_REPRISEARR'));
                                        TF.Post;
                                end;
                        end;
                        TF.EndUpdate;
                        TF.EnableTOM;
                end;
       // end;
end;

procedure TOM_HISTORETENUE.OnNewRecord ;
begin
  Inherited ;
        SetField('PHR_SALARIE',Salarie);
        SetField('PHR_ORDRE',NumOrdre);
        SetField('PHR_DATEDEBUT',Date);
        SetField('PHR_DATEFIN',Date);
        SetField('PHR_DATEPAIEMENT',V_PGI.DateEntree);
        SetField('PHR_EFFECTUE','X');
end ;

procedure TOM_HISTORETENUE.OnLoadRecord ;
var Rubrique : String;
begin
  Inherited ;
  If not (Ecran is TFSaisieList) Then Exit;
        If GetField('PHR_RETENUESAL') = '001' then Rubrique := '7230';
        If GetField('PHR_RETENUESAL') = '002' then Rubrique := '7240';
        If GetField('PHR_RETENUESAL') = '003' then Rubrique := '7250';
        If GetField('PHR_RETENUESAL') = '004' then Rubrique := '7260';
        TFSaisieList(Ecran).RichMessage1.Clear;
        If ExisteSQL('SELECT PHR_SALARIE FROM HISTORETENUE '+
        'LEFT JOIN HISTOBULLETIN ON PHB_DATEDEBUT=PHR_DATEDEBUT AND PHB_DATEFIN=PHR_DATEFIN '+
        'AND PHB_SALARIE=PHR_SALARIE WHERE PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="'+Rubrique+'" '+
        'AND PHB_SALARIE="'+GetField('PHR_SALARIE')+'" AND PHB_DATEDEBUT="'+UsDateTime(GetField('PHR_DATEDEBUT'))+'" '+
        'AND PHB_DATEFIN="'+UsDateTime(GetField('PHR_DATEFIN'))+'" AND PHB_MTREM='+IntToStr(GetField('PHR_MONTANT'))+'') then
        begin
                SetControlEnabled('PHR_DATEDEBUT',False);
                SetControlEnabled('PHR_DATEFIN',False);
                SetControlEnabled('PHR_MONTANT',False);
                SetControlEnabled('PHR_MONTANTMENS',False);
                SetControlEnabled('PHR_DATEPAIEMENT',False);
                SetControlEnabled('PHR_EFFECTUE',False);
                TFSaisieList(Ecran).RichMessage1.Lines.Append('Ce versement correspond au bulletin de paie du '+DateToStr(GetField('PHR_DATEDEBUT'))+ ' au '+DateToStr(GetField('PHR_DATEFIN')));
        end
        else
        begin
                SetControlEnabled('PHR_DATEDEBUT',True);
                SetControlEnabled('PHR_DATEFIN',True);
                SetControlEnabled('PHR_MONTANT',True);
                SetControlEnabled('PHR_MONTANTMENS',True);
                SetControlEnabled('PHR_DATEPAIEMENT',True);
                SetControlEnabled('PHR_EFFECTUE',True);
                If GetField('PHR_EFFECTUE') = 'X' then TFSaisieList(Ecran).RichMessage1.Lines.Append('Ce versement correspond à un remboursement partiel')
                else TFSaisieList(Ecran).RichMessage1.Lines.Append('Ce versement sera pris en compte sur le bulletin du '+DateToStr(GetField('PHR_DATEDEBUT'))+ ' au '+DateToStr(GetField('PHR_DATEFIN')));
        end;
end ;

procedure TOM_HISTORETENUE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_HISTORETENUE.OnArgument ( S: String ) ;
var Q : TQuery;
    LibRetenue : String;
    BImprime : TToolBarButton97;
begin
  Inherited ;
        Salarie := ReadTokenPipe(S,';');
        NumOrdre := ReadTokenPipe(S,';');
        LibRetenue := '';
        If NumOrdre <> '' then
        begin
          Q := OpenSQL('SELECT PRE_LIBELLE,PRE_MONTANTTOT FROM RETENUESALAIRE WHERE PRE_SALARIE="'+Salarie+'" AND PRE_ORDRE='+NumOrdre+'',True);
          If Not Q.Eof then
          begin
                  LibRetenue := Q.FindField('PRE_LIBELLE').AsString;
                  MtTotalRetenue := Q.FindField('PRE_MONTANTTOT').AsFloat;
          end
          else
          begin
                  PGIBox('Aucune retenue n''existe pour ce salarié',Ecran.Caption);
          end;
          Ferme(Q);
        end;
        If Ecran is TFSaisieList Then
        begin
          TFSaisieList(Ecran).Caption := 'Salarié : '+Salarie+' '+RechDom('PGSALARIE',Salarie,False)+', Retenue : '+LibRetenue;
          SetControlEnabled('PHR_SALARIE',False);
          TF  :=  TFSaisieList(Ecran).LeFiltre;
          BImprime := TToolBarButton97(GetControl('BImprimer'));
          If BImprime <> Nil then BImprime.OnClick := BEtatClick;
        end;

end ;

procedure TOM_HISTORETENUE.BEtatClick (Sender : TObject);
VAR St :String;
   Pages : TPageControl;
   StPages : String;
begin
        Pages := TPageControl(GetControl('PAGES'));
                St := 'SELECT PRE_SALARIE,PRE_ORDRE,PHR_SALARIE,PHR_MONTANT,PHR_ARRIERE,PHR_CUMULVERSE,PHR_MONTANTMENS,PHR_DATEPAIEMENT,PHR_ORDRE,PHR_DATEDEBUT,PHR_DATEFIN,PHR_REPRISEARR,'+
        'PRE_LIBELLE,PRE_RETENUESAL,PRE_NBMOIS,PRE_MONTANTMENS,PRE_MONTANTTOT,PRE_DATEDEBUT,PRE_DATEFIN,'+
        'ANN_NOMPER,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE,ANN_APPAYS '+
        'FROM RETENUESALAIRE LEFT JOIN HISTORETENUE ON PHR_SALARIE=PRE_SALARIE AND PHR_ORDRE=PRE_ORDRE '+
        'LEFT JOIN ANNUAIRE ON PRE_BENEFRSGU=ANN_GUIDPER '+
        'WHERE PHR_SALARIE="'+GetField('PHR_SALARIE')+'" AND PHR_ORDRE='+IntToStr(GetField('PHR_ORDRE'))+''+
        ' ORDER BY PRE_SALARIE,PRE_ORDRE,PHR_DATEDEBUT';
        {$IFNDEF EAGLCLIENT}
        LanceEtat('E','PRS','PRE',True,False,False,Pages,St,'',False);
        {$ELSE}
        StPages := AglGetCriteres (Pages, FALSE);
        LanceEtat('E','PRS','PRE',True,False,False,Pages,St,'',False,0,StPages);
        {$ENDIF}
end;

Initialization
  registerclasses ( [ TOM_HISTORETENUE ] ) ;
end.

