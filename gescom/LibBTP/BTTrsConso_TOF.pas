{***********UNITE*************************************************
Auteur  ...... : B. LOCATELLI
Créé le ...... : 5/08/2011
Modifié le ... :   /  /
Description .. : Transfert des lignes de conso inter-chantiers.
Mots clefs ... : TOF;BTTrsConso
*****************************************************************}
Unit BTTrsConso_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     EntGC,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HTB97,
     HDB,
     UTOB,
     FE_Main,
     forms,
     vierge,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AffaireUtil,
     FactGrp,
     ParamSoc,
     utofAfBaseCodeAffaire,
     FactTOB,
     FactUtil,
     CalcOLEGenericBTP,
     UTOF,uEntCommun, AGLInit ;


Type
  TOF_BTTrsConso = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
//    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;

  private

    //Variables correspondant aux objets de la liste
    BOuvrir1   	 : TToolBarButton97;
    BSELECTAFF1  : TToolBarButton97;
    BEFFACEAFF   : TToolBarButton97;

    BCO_AFFAIRE  : THEdit;
    BCO_AFFAIRE0 : THEdit;
    BCO_AFFAIRE1 : THEdit;
    BCO_AFFAIRE2 : THEdit;
    BCO_AFFAIRE3 : THEdit;
    BCO_AVENANT  : THEdit;

    Affaire0		 : String;

  	//Procedure propre à l'appli
    Procedure TransfertConsoLigne(TobConso : Tob; Affaire : String);
    Procedure ControleChamp(Champ, Valeur: String);
    Procedure ControleCritere(Critere: String);

    //procedure liée à un objet de la liste
    procedure BEffaceClick(Sender: TOBJect);
    procedure BSelectAffClick(Sender: TOBJect);
    procedure TransfertConso(Sender: TOBJect);


  end ;

Implementation
uses BTGENODANAL_TOF;
(*
procedure TOF_BTTrsConso.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('BCO_AFFAIRE'));
Aff0 := THEdit(GetControl('BCO_AFFAIRE0'));
Aff1:=THEdit(GetControl('BCO_AFFAIRE1'));
Aff2:=THEdit(GetControl('BCO_AFFAIRE2'));
Aff3:=THEdit(GetControl('BCO_AFFAIRE3'));
end;
*)
procedure TOF_BTTrsConso.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTrsConso.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTTrsConso.OnUpdate ;
begin
  Inherited ;
  TFMul(Ecran).Fliste.ClearSelected;
end ;

procedure TOF_BTTrsConso.OnLoad ;
begin
  Inherited ;
  if ThCheckBox(getControl('CBLIVCHA')).State = cbchecked then
  begin
    THEdit(GetControl('XX_WHERE2')).Text := 'AND ((BCO_NATUREPIECEG="LBT") OR (BCO_NATUREPIECEG="FAC"))';
  end else if ThCheckBox(getControl('CBLIVCHA')).State = cbUnchecked  then
  begin
    THEdit(GetControl('XX_WHERE2')).Text := 'AND ((BCO_NATUREPIECEG="") OR (BCO_NATUREPIECEG="FAC"))';
  end else
  begin
    THEdit(GetControl('XX_WHERE2')).Text := 'AND ((BCO_NATUREPIECEG="") OR (BCO_NATUREPIECEG="LBT") OR (BCO_NATUREPIECEG="FAC"))';
  end;
  if ((GetControlText('BCO_AFFAIRE1') <> '') Or (GetControlText('BCO_AFFAIRE2') <> '') Or (GetControlText('BCO_AFFAIRE3') <> '')) then
    SetControltext('XX_WHERE','')
  else
    SetControltext('XX_WHERE','1=2')
end ;

procedure TOF_BTTrsConso.OnArgument (S : String ) ;
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
begin
  fMulDeTraitement := True;
  Inherited ;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
      begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
        begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	ControleChamp(champ, valeur);
				end
      end;
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(S)));
    end;

    Affaire0 := 'A'; // On ne traite que les affaires de type chantier pour l'instant

    BOuvrir1 := TToolbarButton97(ecran.FindComponent('BOuvrir1'));
		BOuvrir1.OnClick := TransfertConso;
    BOuvrir1.hint := 'Transférer les lignes de consommations';

    BEffaceAFF := TToolbarButton97(ecran.FindComponent('BEffaceAff'));
    BEffaceAFF.OnClick := BEffaceClick;

	  BSelectAFF1  := TToolBarButton97(ecran.FindComponent('BSELECTAFF1'));
    BSelectAFF1.OnClick := BSelectAffClick;

  	BCO_AFFAIRE  := THEdit(GetControl('BCO_AFFAIRE'));
  	BCO_AFFAIRE0 := THEdit(GetControl('BCO_AFFAIRE0'));
  	BCO_AFFAIRE1 := THEdit(GetControl('BCO_AFFAIRE1'));
  	BCO_AFFAIRE2 := THEdit(GetControl('BCO_AFFAIRE2'));
  	BCO_AFFAIRE3 := THEdit(GetControl('BCO_AFFAIRE3'));
  	BCO_AVENANT  := THEdit(GetControl('AFF_AVENANT'));

    BCO_AFFAIRE0.text := Affaire0;

   	ChargeCleAffaire(BCO_AFFAIRE0, BCO_AFFAIRE1, BCO_AFFAIRE2, BCO_AFFAIRE3, BCO_AVENANT, BSelectAFF1, TaModif, BCO_AFFAIRE.Text, false);

end ;

Procedure TOF_BTTrsConso.ControleCritere(Critere : String);
Begin

end;

Procedure TOF_BTTrsConso.ControleChamp(Champ : String;Valeur : String);
Begin

end;

procedure TOF_BTTrsConso.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTTrsConso.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTrsConso.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTTrsConso.BSelectAffClick(Sender : TOBJect);
var Aff0				 : String;
    Aff1				 : String;
    Aff2				 : String;
    Aff3				 : String;
    Aff4				 : String;
    Affaire 		 : string;
    EtatAffaire	 : String;
    bProposition : Boolean;
begin
  Aff0:=affaire0;
  Aff1:='';
  Aff2:='';
  Aff3:='';
  Aff4:='';
  Affaire :='';
  bProposition := False;

  if BCO_AFFAIRE <> Nil then Affaire := BCO_AFFAIRE.Text;

  if not GetAffaireEnteteSt(nil,BCO_AFFAIRE1,BCO_AFFAIRE2,BCO_AFFAIRE3,BCO_AVENANT,Nil,affaire,bProposition,false,false,true,false,'',false,true,false,EtatAffaire ) then Affaire := '';
 
  BCO_AFFAIRE.text := Affaire;

  BTPCodeAffaireDecoupe(affaire,Aff0, Aff1,Aff2,Aff3,aff4,taCreat,false);

  BCO_AFFAIRE0.text := Aff0;
  BCO_AFFAIRE1.text := Aff1;
  BCO_AFFAIRE2.text := Aff2;
  BCO_AFFAIRE3.text := Aff3;
end;

Procedure TOF_BTTrsConso.BEffaceClick(Sender : TOBJect);
Begin
	BCO_AFFAIRE.text  := '';
	BCO_AFFAIRE1.Text := '';
 	BCO_AFFAIRE2.Text := '';
 	BCO_AFFAIRE3.Text := '';
end;

Procedure TOF_BTTrsConso.TransfertConso(Sender : TOBJect);
var req      : String;
    Affaire  : String;
		TobConso, TOBR : Tob;
    NumConso : Double;
    DateConso: TDateTime;
    Indice   : Integer;
    F 			 : TFMul;
    i 			 : integer;
    QQ			 : TQuery;
    QConso	 : TQuery;
    L 			 : THDBGrid;

begin
  Inherited ;

	F:=TFMul(Ecran);
  //controle si au moins un éléments sélectionné
	if (F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  	 begin
	   PGIInfo('Aucun élément sélectionné','');
  	 exit;
   	 end;

  // Sélection de l'affaire destination
  TOBR := TOB.Create ('UNE TOB',nil,-1);
  TOBR.AddChampSupValeur ('RETOUR','');
 	TheTOB := TOBR;
 	AGLLanceFiche('BTP','BTCHOIXAFFAIRE','','','');
  //

  TheTOB := nil;
  Affaire := TOBR.getValue('RETOUR');
 	TOBR.free;
  if Affaire = '' then
  begin
    PGIInfo('Vous n''avez pas indiqué d''affaire destination','');
    Exit;
  end;
  if not ExisteSql('SELECT 1 FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"') then
  begin
    PGIInfo('Merci de renseigner une affaire existante...','');
    Exit;
  end;
  //Demande de confirmation du traitement
	if PGIAsk('Confirmez-vous le traitement ?','') <> mrYes then exit ;

  TOBConso  := Tob.create('Les consos', nil, -1);

	L:= F.FListe;

  SourisSablier;

  TRY
	if L.AllSelected then
  	 begin
  	 QQ:=F.Q;
     QQ.First;
     while not QQ.EOF do
     	Begin
      TOBConso.InitValeurs;
      //chargement clé
      DateConso:= QQ.findfield('BCO_DATEMOUV').asDateTime;
      NumConso := QQ.findfield('BCO_NUMMOUV').AsFloat;
      Indice   := QQ.findfield('BCO_INDICE').asInteger;
      Req := 'SELECT * FROM CONSOMMATIONS '+
						 'WHERE BCO_DATEMOUV="'+USDATETIME(DateConso)+'" AND '+
             'BCO_NUMMOUV='+FloatToStr(NumConso)+' AND '+
             'BCO_INDICE='+IntToStr(Indice);
      Qconso := OpenSQL(Req,true);
      if not QConso.eof then
      begin
        TOBConso.LoadDetailDB('CONSOMMATIONS','','',Qconso,true);
      end;
      Ferme (QConso);
      //
     	QQ.next;
      end;
     end
  else
     begin
     for i:=0 to L.NbSelected-1 do
         begin
	       L.GotoLeBookmark(i);
         NumConso := TFMul(F).Fliste.datasource.dataset.FindField('BCO_NUMMOUV').AsFloat;
         DateConso:= TFMul(F).Fliste.datasource.dataset.FindField('BCO_DATEMOUV').asDateTime;
    	   Indice   := TFMul(F).Fliste.datasource.dataset.FindField('BCO_INDICE').asInteger;
		     Req := 'SELECT * FROM CONSOMMATIONS '+
						    'WHERE BCO_DATEMOUV="'+USDATETIME(DateConso)+'" AND '+
                'BCO_NUMMOUV='+FloatToStr(NumConso)+' AND '+
                'BCO_INDICE='+IntToStr(Indice);
         Qconso := OpenSQL(Req,true);
         if not QConso.eof then
         		begin
            TOBConso.LoadDetailDB('CONSOMMATIONS','','',Qconso,true);
        	  end;
         Ferme (QConso);
         end;
	   end;
  //
  TransfertConsoLigne(TobConso, Affaire);

	FINALLY
    SourisNormale;
    F.BChercheClick(ecran);
    TobConso.Free;
  END;

end;

Procedure TOF_BTTrsConso.TransfertConsoLigne(TobConso : Tob; Affaire : String);
Var Part0, Part1, Part2, Part3, CodeAvenant : String;
    i : Integer;
Begin

  BTPCodeAffaireDecoupe(affaire, Part0,Part1, Part2, Part3,CodeAvenant, TaModif, False);
  For i := 0 to TobConso.detail.count-1 do
  Begin
    TobConso.detail[i].Putvalue('BCO_AFFAIRE',Affaire);
    TobConso.detail[i].Putvalue('BCO_AFFAIRE0',Part0);
    TobConso.detail[i].Putvalue('BCO_AFFAIRE1',Part1);
    TobConso.detail[i].Putvalue('BCO_AFFAIRE2',Part2);
    TobConso.detail[i].Putvalue('BCO_AFFAIRE3',Part3);
    UpdateStatusMoisOD (TobConso);
  end;
  TOBConso.UpdateDB();
end;

Initialization
  registerclasses ( [ TOF_BTTrsConso ] ) ;
end.

