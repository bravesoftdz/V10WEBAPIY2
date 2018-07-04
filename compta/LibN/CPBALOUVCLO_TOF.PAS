{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALOUVCLO ()
Mots clefs ... : TOF;CPBALOUVCLO
*****************************************************************}
Unit CPBALOUVCLO_TOF ;

Interface

Uses StdCtrls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_main,
     QRS1,
{$else}
     eMul,
     uTob,
     Maineagl,
     eQRS1,
{$ENDIF}
     sysutils,
     Ent1,
     HCtrls,
     HEnt1,
     UTOF,
     uLibExercice,  // CExerciceVersRelatif
     ParamSoc,      // GetParamSocSecur
     HTB97,
     CPBALGEN_TOF;

  procedure CPLanceFiche_CPBALOUVCLO(TypeEdition: Char; Argument: String='');
  procedure LanceEditionLegale_BALGENE(Exercice: String; DateDeb, DateFin: TDateTime;
                                       Apercu, Couleur: Boolean);

Type
  TOF_CPBALOUVCLO = Class (TOF_CPBALGEN)
  private
    FTypeEdition: Char;
    FDateDeb:     TDateTime;
    FDateFin:     TDateTime;
    FAvecComparatif: Boolean;         {Balance avec comparatif}

    TitreEtat:    THEdit;
  public
    procedure InitZones; override;
    procedure InitZonesLegale(Argument: String);
    procedure ParamTitresCols; override;
    function  GetConditionSQLCpt:    String ; override;
    function  RecupWhereQualifPiece: String ; override;
    function  GetConditionSQLPeriode(vInPer: Integer): String; override;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure CPLanceFiche_CPBALOUVCLO(TypeEdition: Char; Argument: String='');
begin
  AGLLanceFiche('CP','CPBALOUVCLO_QR1','','',TypeEdition+';'+Argument);
end;

procedure LanceEditionLegale_BALGENE(Exercice: String; DateDeb, DateFin: TDateTime; Apercu, Couleur: Boolean);
var Argument: String;
begin
  Argument := Exercice+';'+DateToStr(DateDeb)+';'+DateToStr(DateFin)+';';

  if Apercu then
    Argument := Argument + 'X;'
  else
    Argument := Argument + '-;';

  if Couleur then
    Argument := Argument + 'X;' 
  else
    Argument := Argument + '-;';

  CPLanceFiche_CPBALOUVCLO('L', Argument);
end;

procedure TOF_CPBALOUVCLO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALOUVCLO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPBALOUVCLO.OnUpdate ;
var
  lStReq: String;
begin
  Inherited ;
  (* FQ 21348 BVE 06.09.07
  if ((FTypeEdition in ['O','L']) and FAvecComparatif) or  {O: Ouverture - F: Fermeture}
     (FTypeEdition = 'F') then  *)
  if ((FTypeEdition in ['O','L']) and FAvecComparatif) then
    begin
    // >> Maj des soldes débiteurs
    lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = (CED_DEBIT1-CED_CREDIT1), CED_CREDIT1=0 '
            + ' WHERE ( CED_DEBIT1 >= CED_CREDIT1 ) AND CED_USER="'+V_PGI.User+'"' ;
    ExecuteSQL( lStReq ) ;
    // >> Maj des soldes créditeurs
    lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = (CED_CREDIT1-CED_DEBIT1), CED_DEBIT1 = 0 '
            + ' WHERE ( CED_DEBIT1 < CED_CREDIT1 ) AND CED_USER="'+V_PGI.User+'"' ;
    ExecuteSQL( lStReq ) ;
    end;

  (* if ((FTypeEdition in ['O', 'F','L']) and FAvecComparatif) then  {O: Ouverture - F: Fermeture} *)
  if ((FTypeEdition in ['O','L']) and FAvecComparatif) then
    begin
    // >> Maj des soldes débiteurs
    lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = (CED_DEBIT2-CED_CREDIT2), CED_CREDIT2=0 '
            + ' WHERE ( CED_DEBIT2 >= CED_CREDIT2 ) AND CED_USER="'+V_PGI.User+'"' ;
    ExecuteSQL( lStReq ) ;
    // >> Maj des soldes créditeurs
    lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = (CED_CREDIT2-CED_DEBIT2), CED_DEBIT2 = 0 '
            + ' WHERE ( CED_DEBIT2 < CED_CREDIT2 ) AND CED_USER="'+V_PGI.User+'"' ;
    ExecuteSQL( lStReq ) ;
    end;

  {Remet les valeurs initiales de l'état}
  (*if (FTypeEdition = 'F') and (not FAvecComparatif) then
    AvecComparatif.Checked := False;*)
  { END FQ 21348 }
  DateComptaDe.Text := DateToStr(FDateDeb);
  DateComptaA.Text  := DateToStr(FDateFin);
  { FQ 21347 BVE 04.09.07 }
  lStReq := 'DELETE FROM ' + GetTablePourBase('CEDTBALANCE') + ' WHERE CED_DEBIT1 = 0 AND CED_CREDIT1 = 0 AND CED_DEBIT2 = 0 AND CED_CREDIT2 = 0 AND CED_USER = "' + V_PGI.User + '"';
  ExecuteSQL( lStReq ) ;
  { END FQ 21347 }
end ;

procedure TOF_CPBALOUVCLO.OnLoad ;
begin
  FDateDeb        := StrToDate(DateComptaDe.Text);
  FDateFin        := StrToDate(DateComptaA.Text);
  FAvecComparatif := AvecComparatif.Checked;

  TFQRS1(Ecran).NatureEtat := 'CLO';
  case FTypeEdition of
    'O' : begin      {Ouverture}
          if FAvecComparatif then
            begin
            { FQ 21350 BVE 07.09.07
            TFQRS1(Ecran).CodeEtat   := 'COC';}
            TFQRS1(Ecran).CodeEtat   := 'CFS';
            { END FQ 21350 }
            end
          else
            begin
            TFQRS1(Ecran).CodeEtat   := 'COS';
            {Il faut éditer les écritures générées le premier jour de l'exercice}
            DateComptaA.Text := DateComptaDe.Text;
            end;
          end;
    'F' : begin      {Fermeture}
          if FAvecComparatif then
            begin               
            { FQ 21350 BVE 07.09.07
            TFQRS1(Ecran).CodeEtat   := 'COC'; }
            TFQRS1(Ecran).CodeEtat   := 'CFS';
            { END FQ 21350 }
            end
          else
            begin
            {Il faut éditer les écritures générées le premier jour de l'exercice}
            { FQ 19406 BVE 01.06.07
            DateComptaDe.Text := DateToStr(StrToDate(DateComptaDe.Text) - 1);
            DateComptaA.Text := DateComptaDe.Text;
             END FQ 19406 }
            TFQRS1(Ecran).CodeEtat   := 'CFS';
            end;
          end;
    'L' : begin      {Legale}
            {Il faut éditer les écritures générées le premier jour de l'exercice}
            DateComptaA.Text := DateComptaDe.Text;
            TFQRS1(Ecran).CodeEtat   := 'CFS';
          end;
  end;

  (* FQ 21348 BVE 06.09.07
  if (FTypeEdition = 'F') and (not FAvecComparatif) then begin   {Fermeture uniquement}
    AvecComparatif.Checked := True;
    ComparType.ItemIndex := 0;
    ComparExo.Value      := Exercice.Value;
    end;
   END FQ 21348 *) 
  Inherited ;
end ;

procedure TOF_CPBALOUVCLO.OnArgument (S : String ) ;
var
  Argument: String;
begin
  Argument        := S;
  FTypeEdition    := ReadTokenSt(Argument)[1];
  FAvecComparatif := False;

  // 19/07/2007 : PATCH à mettre dans la fiche en V9 !!!!!!!!!!!
  { BVE 28.09.07 : GCYCLEREVISION est un THEDit desormais
  THMultiValComboBox( GetControl('GCYCLEREVISION', True)).DataType := 'CREVCYCLEACTIF' ;}

  Inherited OnArgument(Argument);

  TitreEtat := THEdit(Getcontrol('TITREETAT', true));

  case FTypeEdition of
    'O' : begin      {Ouverture}
          Ecran.Caption     := TraduireMemoire('Balance d''ouverture');
          TFQRS1(Ecran).FNomFiltre:='OUVERTURE';
          Ecran.HelpContext := 7760000;
          end;
    'F' : begin      {Fermeture}
          Ecran.Caption     := TraduireMemoire('Balance de clôture');
          TFQRS1(Ecran).FNomFiltre:= 'FERMETURE';
          Ecran.HelpContext := 7769000;
          end;
    'L' : begin      {Légale}
          Ecran.Caption     := TraduireMemoire('Balance générale');
          TFQRS1(Ecran).FNomFiltre:= 'LEGALE';
          //Ecran.HelpContext := ;
          end;
  end;
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;

  if FTypeEdition = 'L' then    {Lance l'édition}
    begin
    InitZonesLegale(Argument);
    (GetControl('BAGRANDIR', True) as TToolbarButton97).Click;
    (GetControl('BVALIDER', True)  as TToolbarButton97).Click;
    end;

  { FQ 21351 BVE 04.09.07 }
  TFQRS1(Ecran).ParamEtat := False;
  TFQRS1(Ecran).ChoixEtat := False;
  { END FQ 21351 }
  { FQ 21348 BVE 13.09.07 }
  SetControlEnabled('BILAN',False);
  { END FQ 21348 }
end ;

procedure TOF_CPBALOUVCLO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPBALOUVCLO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPBALOUVCLO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPBALOUVCLO.InitZones;
begin
  inherited;
  case FTypeEdition of
    'O' : begin      {Ouverture}
          end;
    'F' : begin      {Fermeture}
          if VH^.Precedent.Code <> '' then
            Exercice.Value := CExerciceVersRelatif(VH^.Precedent.Code);
          end;
  end;
end;

procedure TOF_CPBALOUVCLO.InitZonesLegale(Argument: String);
begin
  Exercice.Value    := CExerciceVersRelatif(ReadTokenSt(Argument));
  DateComptaDe.Text := ReadTokenSt(Argument);
  DateComptaA.Text  := ReadTokenSt(Argument);
  SetControlChecked('FAPERCU',  ReadTokenSt(Argument) = 'X');
  SetControlChecked('FCOULEUR', ReadTokenSt(Argument) = 'X');

  AvecComparatif.Checked := True;
  ComparType.ItemIndex   := 0;
  ComparExo.Value   := Exercice.Value;
  ComparDateDe.text := DateComptaDe.Text;
  ComparDateA.text  := DateComptaA.Text;
  QualifPiece.Value :='N';

  ModeSelect.ItemIndex := 0;
  SetControlProperty('DEVISE', 'DATATYPE', TraduireMemoire('<<Tous>>') ) ;
  SetControlVisible('FFILTRES',  False);
  SetControlVisible('BFILTRE',   False);
  SetControlVisible('BAGRANDIR', False);
  SetControlVisible('BVALIDER',  False);
end;

procedure TOF_CPBALOUVCLO.ParamTitresCols;
begin
  inherited;
  case FTypeEdition of
    'O',
    'F' : begin      {Ouverture Fermeture Lek 24/05/06 FQ 18093}
          { FQ 21348 BVE 06.09.07
          SetControlText('TitreCol3', TraduireMemoire('Total'));
          SetControlText('TitreCol4', TraduireMemoire('Ecart'));
            END FQ 21348 }
          end;
    'L' : begin      {Légale}
          SetControlText('TitreCol1','au ' + DateToStr(FdateDeb - 1)) ;
          SetControlText('TitreCol2','du ' + DateToStr(FdateDeb) + ' au ' + DateToStr(FDateFin)) ;
          SetControlText('TitreCol3','au ' + DateToStr(FdateFin)) ;
          SetControlText('TitreCol4','Solde au ' + DateToStr(FdateFin)) ;
          end;
  end;
end;

function TOF_CPBALOUVCLO.RecupWhereQualifPiece: String;
begin
  case FTypeEdition of
    'O','L' : begin      {Ouverture}
              Result := ' AND E_QUALIFPIECE="N"';
              end;
    'F' : begin      {Fermeture}
          { FQ 19406 BVE 01.06.07
          Result := ' AND (E_QUALIFPIECE="N" OR E_QUALIFPIECE="C")';
          }           
          Result := ' AND E_QUALIFPIECE="C" AND E_QUALIFPIECE<>"I"'; // FQ 21396 BVE 10.09.06
          { END FQ 19406 }
          end;
  end;
end;

function TOF_CPBALOUVCLO.GetConditionSQLCpt: String;
var
  ResC,BilC,PerC,BenC,BilO,BenO,PerO : String ;
begin
  Result := inherited GetConditionSQLCpt;
  case FTypeEdition of
    'O' : begin      {Ouverture}
          if not FAvecComparatif then
            Result := Result + ' AND '+
             '( (G_TOTDEBE<>0 OR G_TOTCREE<>0 OR G_TOTDEBANO<>0 OR G_TOTCREANO<>0 )'+
             ' OR (EXISTS (SELECT E_GENERAL,E_EXERCICE,E_DATECOMPTABLE'+
                          ' FROM ECRITURE'+
                          ' WHERE E_GENERAL=G_GENERAL'+
                          ' AND E_EXERCICE="'+CRelatifVersExercice(Exercice.Value)+'"'+
                          ' AND E_DATECOMPTABLE>="'+USDATETIME(FDateDeb)+'" AND E_DATECOMPTABLE<="'+USDATETIME(FDateFin)+'"'+
                          ' AND E_QUALIFPIECE="N" ))' +
            ')';
          end;
    'F' : begin      {Fermeture}
          if (not FAvecComparatif) then begin
            BilC:=GetParamSocSecur('SO_FERMEBIL','') ;
            ResC:=GetParamSocSecur('SO_RESULTAT','') ;
            PerC:=GetParamSocSecur('SO_FERMEPERTE','') ;
            BenC:=GetParamSocSecur('SO_FERMEBEN','') ;
            BilO:=GetParamSocSecur('SO_OUVREBIL','') ;
            PerO:=GetParamSocSecur('SO_OUVREPERTE','') ;
            BenO:=GetParamSocSecur('SO_OUVREBEN','') ;

            Result := Result + ' AND '+
            '( ((G_TOTDEBS<>0 OR G_TOTCRES<>0 OR G_TOTDEBANON1<>0 OR G_TOTCREANON1<>0)'+
              ' OR (EXISTS (SELECT E_GENERAL,E_EXERCICE,E_DATECOMPTABLE'+
                           ' FROM ECRITURE'+
                           ' WHERE E_GENERAL=G_GENERAL'+
                           ' AND E_EXERCICE="'+CRelatifVersExercice(Exercice.Value)+'"'+
                           ' AND E_DATECOMPTABLE>="'+USDATETIME(FDateDeb)+'" AND E_DATECOMPTABLE<="'+USDATETIME(FDateFin)+'"'+
                           ' AND E_QUALIFPIECE="N" )) )'+
             ' OR (G_GENERAL="'+BilC+'")'+
             ' OR (G_GENERAL="'+ResC+'")'+
             ' OR (G_GENERAL="'+PerC+'")'+
             ' OR (G_GENERAL="'+BenC+'")'+
             ' OR (G_GENERAL="'+BilO+'")'+
             ' OR (G_GENERAL="'+PerO+'")'+
             ' OR (G_GENERAL="'+BenO+'")'+
            ')';
          end;
          if GetControlText('BILAN')='-' then
            Result := Result + ' AND G_NATUREGENE<>"CHA"' +
                               ' AND G_NATUREGENE<>"PRO"';
          end;
  end;
end;

function TOF_CPBALOUVCLO.GetConditionSQLPeriode(vInPer: Integer): String;
var Exo : integer;
begin
  result := '';
  Case vInPer of
      // période principale (exo + date)
      1 :
      begin
         if (FTypeEdition = 'F') and not(FAvecComparatif) then
         begin
            { FQ 21348 BVE 06.09.07
            Exo := Exercice.ItemIndex - 1; }
            Exo := Exercice.ItemIndex;
            { END FQ 21348 }
            if Exo >= 0 then
               result := ' E_EXERCICE ="' + CRelatifVersExercice(Exercice.Values[Exo]) + '"';
            result := result + ' AND E_DATECOMPTABLE BETWEEN "' + USDateTime( DateDebutPer ) +
                      '" AND "' + USDateTime( DateFinPer ) + '"';
         end
         else
            result := ' E_EXERCICE ="' + CRelatifVersExercice(Exercice.Value) + '" AND ' +
                      ' E_DATECOMPTABLE BETWEEN "' + USDateTime( DateDebutPer ) +
                      '" AND "'     + USDateTime( DateFinPer ) + '"' ;

      end;
      // Période optionnelle
      2 : result := ' E_EXERCICE ="' + CRelatifVersExercice(ComparExo.Value) + '" AND ' +
                    ' E_DATECOMPTABLE BETWEEN "' + USDateTime( DateDebutCompar ) +
                    '" AND "'     + USDateTime( DateFinCompar ) + '"' ;

      end ;

  if ((FTypeEdition='O') and (not FAvecComparatif)) or (FTypeEdition='L') then begin
    if vInPer=1 then
      Result := Result + ' AND E_ECRANOUVEAU="OAN"'
    else
      Result := Result + ' AND E_ECRANOUVEAU<>"OAN"';
  end;
  if (FTypeEdition='F') and (not FAvecComparatif) then begin
    if vInPer<>1 then
      Result := Result + ' AND E_ECRANOUVEAU<>"OAN"';
  end;
end;

Initialization
  registerclasses ( [ TOF_CPBALOUVCLO ] ) ;
end.

