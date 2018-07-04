{***********UNITE*************************************************
Auteur  ...... :  MCD
Créé le ...... : 01/02/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFAUGMEN(), augmentation affaire
Mots clefs ... : TOF;AFAUGMEN
*****************************************************************}
Unit UTOFAFAUGMEN ;

Interface

Uses
{$IFDEF EAGLCLIENT}
   eMul,
   Maineagl,
{$ELSE}
   Mul,
   AssistCreationAffaire,
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
   HDB,
   FE_Main,
{$ENDIF}
   Controls,
   Classes,
   forms,
   sysutils,
   HCtrls,
   Hstatus,
   HEnt1,
   M3FP,
   HMsgBox,
   UTOF,
   Utob,Dicobtp,
   SaisUtil,
   Factcalc,
   AffaireUtil,
   FactComm,
   Ent1,
   UtilPgi,
   Factutil,
   TarifUtil,
   uTofAfBaseCodeAffaire,
   TraducAffaire,
   AffaireDuplic,
   AffaireModifPiece,
   Utom,
   UtilMulTrt,
   StdCtrls,
   EntGC;

Type
  TOF_AFAUGMEN = Class (TOF_AFBASECODEAFFAIRE)
     procedure Augmen_Mission;
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override ;
     procedure OnUpdate                 ; override ;
     Procedure OnArgument(stArgument : String ); override ;
     procedure PourExit (Sender : Tobject);
     //
     function TraitAugMis : boolean;

  private

     StNonresilie : string;
     TobAffModele	: TOB;
     TobFactaff 	: TOB;
     TobAffaireNew: TOB;
     TobFactaffNew: TOB;
     TOM_AFF:TOM;
     Totalech : double;
     RecupTotaux : string;
     Statut		 		: String;
      MOISDEB : THvalComboBox;
     pour: ThNumedit;
     ZArrondi: ThValCOmboBox;
     ZAnnul: TCheckBox;
     Annul:Boolean;
     dev : RDEVISE;
     Procedure AlimComplementAffModele;
     procedure ControleChamp(Champ, Valeur: String);
     Procedure ValideAffaireCreee;
     Function MajChampsAffaire :Boolean;
    function ChargeLesAffaires(Mul: TFMul; ChampCpt, TableCpt: string;TobEnregs: TOB; Parle: Boolean): T_ErrTraiteEnreg;
    function TestEnregATraiter(Mul: TFMul; Parle: Boolean): Boolean;
    function VerifNbrPourBlocage(ChampCpt, TableCpt, stWhere: string;
      Parle: Boolean; nbEnregMax: integer): Boolean;
    procedure FinTraiteMul(Mul: TFMul);
    procedure MoisDebChange (Sender : Tobject);
    function MajComplAffNewClient(var MtEcheance: double): Boolean;
  //  procedure CalculPort ;
  end ;
  
var titre: string;
Procedure AFLanceFiche_Mul_AugmentAff(Range:string);

Implementation
uses AffEcheanceUtil;

procedure TOF_AFAUGMEN.OnArgument(stArgument : String );
var x				: integer;
    Critere : String;
    ValMul	: String;
    ChampMul: String;
    CC      : THValComboBox;
begin
	fMulDeTraitement := true;
Inherited;
	FTableName := 'AFFAIRE';
  stNonresilie := ' AND (AFF_DATERESIL="12/31/2099")';

  Critere:=uppercase(ReadTokenSt(stArgument));
  while Critere<>'' do
     begin
     x:=pos('=',Critere);
     if x<>0 then
        begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
        end
     else
        ChampMul := Critere;
     ControleChamp(ChampMul, ValMul);
     Critere:= uppercase(Trim(ReadTokenSt(stArgument)));
     end;

	if statut <> '' then setcontrolText('AFF_STATUTAFFAIRE', Statut);
  SetControlProperty ('AFF_ETATAFFAIRE','Plus', ' AND (CC_LIBRE="BTP" AND CC_CODE<>"TER")');

pour:=THNumEdit(GetControl('ZPOUR')) ;
	Pour.OnExit := PourExit;
ZArrondi:=THValCOmboBox(GetControl('ARRONDI')) ;
zAnnul:= TCheckBox(GetControl('ZANNUAUG')) ;

  MOISDEB := THvalComboBox(GetControl('MOISDEB'));
  MOISDEB.OnChange := MoisDebChange;
  // mcd 11/06/02
If not(ctxScot in V_PGI.PGIContexte)  then
    begin
    SetControlVisible ('TT_MOISCLOTURE' , False);
    SetControlVisible ('T_MOISCLOTURE' , False);
    SetControlVisible ('TT_MOISCLOTURE_' , False);
    SetControlVisible ('T_MOISCLOTURE_' , False);
    end;
  THEdit(GetCOntrol('XX_WHERE')).Text := StNonresilie;
  Ecran.Caption := 'Augmentation des Contrats';
  //
  // gestion Etablissement (BTP)
	CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
	if CC<>Nil then
  begin
  	PositionneEtabUser(CC) ;
    if not VH^.EtablisCpta then
    begin
    	if THLabel(GetControl('TBTB_ETABLISSEMENT')) <> nil then THLabel(GetControl('TBTB_ETABLISSEMENT')).Visible := false;
			CC.visible := false;
    end;
	end;

  // gestion Domaine (BTP)
	CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
	if CC<>Nil then PositionneDomaineUser(CC) ;
  //

  UpdateCaption(Ecran);
  
end;

procedure TOF_AFAUGMEN.PourExit (Sender : Tobject);
begin

	if pour.value <0  then
     SetControLVisible('ZANNUAUG' ,true)
  else
     SetControLVisible('ZANNUAUG' ,False);

end;


procedure TOF_AFAUGMEN.OnUpdate ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end ;

procedure TOF_AFAUGMEN.Augmen_Mission;
var  St,Sttexte : String;
     ret : boolean;
     deb,fin  : String;

begin
  if (pour.value = 0.0) then begin
     PGIInfoAF('Aucune information d''augmentation saisie',titre);
     exit;
     end;
  if GetControlText('Arrondi') ='' then
     begin
     if PgiAsk ('Vous n''avez pas saisi de méthode d''arrondi, confirmez-vous quand même le traitement',ecran.caption) =mrno then exit
     end;
  TobAffModele := TOB.Create('AFFAIRE',Nil,-1);


  St := 'Confirmez vous l''augmentation de ces Affaires.#13#10 Attention toutes les échéances non facturés seront augmentées ';
  If (PGIAskAf(st,titre)<> mrYes) then exit;

  deb := FormatDateTime('dd/mm/yyyy ttttt',NowH);
  Ret := TraitAugMis;
  fin := FormatDateTime('dd/mm/yyyy ttttt',NowH);

  if (not Ret) then StTexte := 'Le traitement ne s''est pas fait en totalité'
  else
  if VH_GC.GCIfDefCEGID then
     StTexte := Format('Le traitement s''est terminé correctement de  %s à %s' ,[deb,fin])
  else
  StTexte := 'Le traitement s''est terminé correctement';
  PGIInfoAf(StTexte,titre);
  TobAffModele.free;
END;

function  TOF_AFAUGMEN.TraitAugMis: boolean;
var  wi: Integer;
     F : TFMul ;
     TobMAff : TOB; // Toutes les affaires sélectionnées
     {NbPiece,} Err: Integer;
     TraitementOK : Boolean;
     QQ : Tquery;
     stSql:string;
     DateDevise : TDateTime;
     dDateDebItv, dDateFinItv : TDateTime;
     MTEcheance : double;
     //
		 MtDoc: double ;
     CalcEcheance,stPeriode: string;
     Interval,NbIt : integer;
     TypeCalcul  : string;
     DateDebut   : TDateTime;
     DateFin     : TDateTime;
     //
begin
  result := true;
//  Err:=0;
  // Gestion d'une Tob pour stocker toutes les affaires à traiter
  // Création de cette tob à partir de la multi-sélection  du multi-critère
  // Ce query est basé sur la liste AFAFFTIERS
  F:=TFMul(Ecran);
  TobMAff := Tob.Create('les affaires',NIL,-1);

  // PA le 29/08/2001 - Fonction de traitement des enreg du mul externalisée
  ChargelesAffaires (F, 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);

  InitMove(TobMaff.Detail.count*2,' ');
  wi := 0;
  while wi < TobMAff.Detail.count  do
      Begin
      MoveCur(False);
      TraitementOK :=True;

  if ( VH_GC.GCIfDefCEGID ) or
     ( (not VH_GC.GCIfDefCEGID) and (BlocageAffaire( 'MAF', TobMAff.Detail[wi].GetValue('AFF_AFFAIRE'), V_PGI.groupe, V_PGI.DateEntree, False, False, false, dDateDebItv, dDateFinItv, nil)<>tbaAucun))
           then  begin
           PGIInfoAf(Format('Affaire %s %s non modifiable',[TobMAff.Detail[wi].getValue('AFF_TIERS'),TobMAff.Detail[wi].getValue('AFF_AFFAIRE')]),Titre);
           end
       else begin
          // on n'a pas besoin de tous les champs, mais d'une grande majorité. fait pour une affaire seulement. on laisse
       stSQL := 'SELECT * From AFFAIRE where AFF_AFFAIRE="'+ TobMAff.Detail[wi].getValue('AFF_AFFAIRE') +'"';
       QQ := OpenSQL (stSQL,True);
       if Not (QQ.EOF) then
           BEGIN
           MTEcheance := 0;
           if  TobAffModele <> Nil then begin  TobAffModele.free; TobAffModele :=Nil;end;
            TobAffModele := TOB.Create('AFFAIRE',Nil,-1);
            TOM_AFF := CreateTom('AFFAIRE',Nil,False, False);

            TobAffModele.SelectDB ('',QQ);
           Dev.Code :=  TobAffModele.GetValue ('AFF_DEVISE');
           if Dev.code = '' then DEV.Code:= V_PGI.DevisePivot;
           GetInfosDevise(DEV) ;
           DateDevise :=  TobAffModele.GetValue ('AFF_DATECREATION');
           DEV.Taux:= GetTaux(DEV.Code,DEV.DateTaux,DateDevise) ;
             // charge mission en cours
          AlimComplementAffModele;
                // on crée la nouvelle affaire
          if  TobAffaireNew <> Nil then BEGIN  TobAffaireNew.Free;  TobAffaireNew := Nil; END;
             TobAffaireNew := Tob.create('AFFAIRE',Nil,-1);
             TobAffaireNew.Dupliquer ( TobAffModele,False,True,False);
          if (TObAffaireNew.Getvalue('AFF_CALCTOTHTGLO')='X') then TobAffaireNew.putValue('AFF_TOTALHTGLODEV',0); // mcd 28/11/00 car si calcul HT à partir eche ne marche pas ..;
            if TraitementOK then
                   begin
                     if Zannul.checked then Annul:=true else Annul:=false;
                     RecupTotaux :='POUR:'+FLoatToStr(pour.value)+';ARRONDI:'+Zarrondi.value+';ANNUL:';
                     If Annul then RecupTotaux:=RecupTotaux+'TRUE' else  RecupTotaux:=RecupTotaux+'False';
                     err := ModifieLaPieceAff (tmpAugment, TobMAff.Detail[wi].getValue('AFF_AFFAIRE'), Nil,False,RecupTotaux);
                         // mcd 22/01/02 attention si err <>0, il doit falloir positionner aussi traitementOk ??
                     if err <> 0 then
                     begin
                        V_PGI.IoError:=oeUnknown;
                        TraitementOk := false;
                     end else
                     begin
                        RecalculTotAffaire(RecupTotaux,TobAffaireNew);
                     end;
                   end;
               //maj table affaire
            if TraitementOK then TraitementOK := MajComplAffNewClient (MTEcheance) ;
            if TraitementOK then TraitementOK := MajChampsAffaire  ;
            if TraitementOk then
            begin
            	(* Faux et archi faux
            	TobAffaireNew.PutValue ('AFF_MONTANTECHEDEV' ,MTEcheance);
              if DEV.Code <> V_PGI.DevisePivot then
              begin
            		TobAffaireNew.PutValue ('AFF_MONTANTECHE' ,DeviseToPivot(MTEcheance,DEV.Taux,Dev.Quotite));
              end else
              begin
            		TobAffaireNew.PutValue ('AFF_MONTANTECHE' ,TobAffaireNew.GetValue ('AFF_MONTANTECHEDEV'));
              end;
              *)
              MtDoc := TOBAffaireNew.getValue('AFF_TOTALHTGLODEV');
              CalcEcheance := TOBAffaireNew.getValue('AFF_DETECHEANCE');
              MtDoc := TOBAffaireNew.getValue('AFF_TOTALHTDEV');
              TypeCalcul :=TOBAffaireNew.getValue('AFF_PERIODICITE') ;
              DateDebut := TOBAffaireNew.getValue('AFF_DATEDEBGENER')  ;
              if (TOBAffaireNew.getValue('AFF_GENERAUTO') = 'CON') and (TOBAffaireNew.getValue('AFF_METHECHEANCE') = 'CIV') then
                // En mode de génération Contrat et méthode d'échéance civile, on place la date de début au début de la période
                // suivant la périodicité : Annuelle, mensuelle, hebdomadaire
                DateDebut := GetDateDebutPeriode (TypeCalcul, DateDebut, TOBAffaireNew.getValue('AFF_INTERVALGENER') , TOBAffaireNew.getValue('AFF_METHECHEANCE')) ;

              DateFin := TOBAffaireNew.getValue('AFF_DATEFINGENER');

              StPeriode  := TOBAffaireNew.getValue('AFF_PERIODICITE');
              Interval   := TOBAffaireNew.getValue('AFF_INTERVALGENER');
              //Modif FV : Gestion des contrats au nombre d'interventions
              if TypeCalcul = 'NBI' then
                NbIt := 1
              else
              	NbIt := EvaluationNbEcheances (TOBAffaireNew.getValue('AFF_AFFAIRE'), TypeCalcul, Interval, DateDebut, DateFin) ;
  						MtEcheance := CalculMtEcheanceContrat (MtDoc,CalcEcheance,stPeriode,Interval,NbIt);
            	TobAffaireNew.PutValue ('AFF_MONTANTECHEDEV' ,MTEcheance);
            	TobAffaireNew.PutValue ('AFF_MONTANTECHE' ,DeviseToPivotEx(MtEcheance,DEV.taux,DEV.Quotite,V_PGI.okdecV));
            end;
            if TraitementOK then
                      if Transactions(ValideAffaireCreee,1)<>oeOk then
                       BEGIN TraitementOK := False; PGIInfoAf('Pb rencontré lors de la validation',Titre);; END;
            if TraitementOK=False then begin
                PGIInfoAf(Format('Client: %s Affaire %s %s non augmentée'
                ,[ Tobaffairenew.getValue('AFF_TIERS'),   Tobaffairenew.getValue('AFF_AFFAIRE1'),
                 Tobaffairenew.getValue('AFF_AFFAIRE2')]),Titre);
                result:=False;
                end;
            MoveCur (False);
            if  TobAffModele <> Nil then begin  TobAffModele.free; TobAffModele :=Nil;end;
            if  Tobfactaff <> Nil then begin  Tobfactaff.Free;  Tobfactaff:=Nil; end;
            if  TobAffaireNew<> Nil then begin  TobAffaireNew.Free; TobAffaireNew:=Nil; end;
            if  TobfactaffNew<>Nil then begin  TobfactaffNew.Free;  TobfactaffNew:=Nil; end;
            if TOM_AFF <> Nil then TOM_AFF.Free;
           end
      else  PGIInfoAf(Format('Affaire %s non trouvée',[TobMAff.Detail[wi].getValue('AFF_AFFAIRE')]),Titre);
      ferme(QQ);
  ////
      end;  //fin si pas affaire bloquer
      inc(wi);
  End;  // fin while  ,boucle sur la sélection

  TobMaff.free;
  FiniMove;
END;

Function TOF_AFAUGMEN.MajChampsAffaire :Boolean;
var mm {,totlig}: double;
    //Nbeche : integer;
    //BaseEche,SaisieContre: Boolean;
    //X : double;
    //sufF, sufE : string;
    zcoef2,zcoef{,zcoefcum} : double;
    zdate : string;
BEGIN
Result := True;
TobAffaireNew.PutValue('AFF_DATEMODIF', V_PGI.DateEntree);
if ( TobAffaireNew.Getvalue('AFF_AFFCOMPLETE')= '-')
   or  ( TobAffaireNew.Getvalue('AFF_CALCTOTHTGLO')<>'X') then
    begin
       // tot affaaire, pas alimenté depuis lignes, il faut l'augmenter
     mm:=   TobAffaireNew.GetValue('AFF_TOTALHTGLODEV');
     mm := CalcPourAugm(pour.value,mm,annul); // mm :=mm +(mm * pour.value);
     mm:=ArrondirPrix(ZArrondi.value,mm);
     mm:=Arrondi(mm,Dev.Decimale);
     TobAffaireNew.PutValue('AFF_TOTALHTGLODEV', mm);
   end;
if (TobAffaireNew.Getvalue('AFF_CALCTOTHTGLO')='X') and
   (TobAffaireNew.GetValue('AFF_TOTALHTGLODEV') =0)  and
   ((TobAffaireNew.GetValue('Aff_GENERAUTO') = 'FOR') or (TobAffaireNew.GetValue('Aff_GENERAUTO') = 'ACT') or (TobAffaireNew.GetValue('Aff_GENERAUTO') = 'MAN'))
    then begin
          // calcul auto mais pas de lignes, on cumul les écéhances
     TobAffaireNew.PutValue('AFF_TOTALHTGLODEV', TotalEch);
   end;

	if (TobAffaireNew.Getvalue('AFF_GENERAUTO')<>'CON') then  //GM si con c'est recalucle dans
  // le onupdate
  Begin
    mm:=  TobAffaireNew.GetValue('AFF_MONTANTECHEDEV');
    mm := CalcPourAugm(pour.value,mm,Annul); //mm :=mm +(mm * pour.value);
    mm:=Arrondi(mm,Dev.Decimale);
    mm:=ArrondirPrix(ZArrondi.value,mm);
    TobAffaireNew.PutValue('AFF_MONTANTECHEDEV', mm);
	End;

  zcoef2 := TobAffaireNew.GetValue('AFF_COEFREVALOCUM');
  
  if (pour.value > 0) then
  begin
    zcoef := 1+pour.value;
  	TobAffaireNew.PutValue('AFF_COEFFREVALO', zcoef);

    if zcoef2 = 0 then zcoef2 := 1;
		zcoef2 :=zcoef2 * zcoef;
    TobAffaireNew.PutValue('AFF_COEFREVALOCUM', zcoef2);
  end
  else
  begin
  	zcoef := 1+ (pour.value * (-1));
  	TobAffaireNew.PutValue('AFF_COEFFREVALO', 1);
    if (zcoef2 <> 0) then
			zcoef2 := zcoef2 / zcoef
    else
    	zcoef2 := 1;
  end;

    if VH_GC.GCIfDefCEGID then
    Begin
      // gm pour les augmentation de CEGID, permet de repérer les contrat augmentés
      // il faudrait mettre une date de dernière augmentation ou qq d'autre
      zdate := datetoStr(V_PGI.DateEntree);
      TobAffaireNew.PutValue('AFF_LIBREAFFA',copy(zdate,7,4));
    End;
    TobAffaireNew.PutValue('AFF_COEFREVALOCUM',Arrondi(zcoef2,6));

END;


Procedure TOF_AFAUGMEN.ValideAffaireCreee;

BEGIN
////// MAJ de l'affaire /////////////
TobAffaireNew.AddChampSupValeur ('CreationAuto','X');
If Not(TOM_AFF.VerifTOB (TobAffaireNew)) then begin V_PGI.IoError:=oeSaisie ; exit; end;
TobAffaireNew.UpdateDBTable;
  // création des écheances
if  TobFactaffNew <> Nil then  TobfactaffNew.UpdateDB(False,False);
END;



Function TOF_AFAUGMEN.MajComplAffNewClient (var MtEcheance : double) : Boolean;
var i : integer;
    TobDet : TOB;
 //   SufF,SufE : string;
    mm : double;
    saisiecontre:boolean;
BEGIN
Result := true;
TotalEch:=0;
// Traitement de factaff
  if  TobFactaffNew <> Nil then BEGIN  TobFactaffNew.Free;  TobFactaffNew := Nil; END;
  if  TobFactaff <> Nil then
     BEGIN
    TobFactaffNew := TOB.Create('echeance des missions',Nil,-1);
    TobFactaffNew.dupliquer ( TobFactaff, True, True, False);
   for i := 0 to  TobFactaffNew.detail.count-1 do
       BEGIN
       TobDet :=  TobFactaffNew.detail[i];
       if (tobdet.getvalue('AFA_ECHEFACT') = '-') then
       begin
         mm:=  Tobdet.GetValue('AFA_MONTANTECHEDEV');
         mm := CalcPourAugm(pour.value,mm,Annul); //mm :=mm +(mm * pour.value);
         mm:=ArrondirPrix(ZArrondi.value,mm);
         mm:=Arrondi(mm,Dev.Decimale);
         Tobdet.PutValue('AFA_MONTANTECHEDEV', mm); MTEcheance := mm;
         TotalEch:=TotalEch + mm;
         if (DEV.Code <> V_PGI.DevisePivot) then
            BEGIN
            // mcd 13/02/03 if VH^.TenueEuro then BEGIN SufF:='CON' ; SufE:='' ; END else BEGIN SufF:='' ; SufE:='CON' ; END ;
            Mm:=TobDet.GetValue('AFA_MONTANTECHEDEV');
            // TobDet.PutValue('AFA_MONTANTECHE'+SufF,DeviseToFranc(Mm,DEV.Taux, DEV.Quotite)) ;
            // TobDet.PutValue('AFA_MONTANTECHE'+SufE,DeviseToEuro (Mm,DEV.Taux, DEV.Quotite)) ;
            TobDet.PutValue('AFA_MONTANTECHE',DeviseToEuro (Mm,DEV.Taux, DEV.Quotite)) ;
            END
        Else
            BEGIN
            if (  TobAffModele.getValue('AFF_SAISIECONTRE')='X') then saisiecontre := TRUE else SaisieCOntre:=False;
            if SaisieContre=False then
               BEGIN
                TobDet.PutValue ('AFA_MONTANTECHE', TobDet.GetValue ('AFA_MONTANTECHEDEV'));
                END
            else
               BEGIN
               mm:= TobDet.GetValue('AFA_MONTANTECHEDEV');
               TobDet.PutValue ('AFA_MONTANTECHE',ConvertSaisieEF(mm ,SaisieContre)) ;
               END;
            END;
        end else   TotalEch:=TotalEch + Tobdet.GetValue('AFA_MONTANTECHEDEV'); // ech déjà facturée
       END;
     END;

END;


Procedure TOF_AFAUGMEN.AlimComplementAffModele;
Var Q : TQuery;
    stSQL : String;
    chp : string; //GM
BEGIN
     //alimente tob factaff
chp:= 'AFA_DATEECHE,AFA_AFFAIRE,AFA_NUMECHE,AFA_TIERS,AFA_ECHEFACT';
//chp:=chp+',AFA_MONTANTECHE,AFA_MONTANTECHECON,AFA_MONTANTECHEDEV,AFA_TYPECHE';
chp:=chp+',AFA_MONTANTECHE,AFA_MONTANTECHEDEV,AFA_TYPECHE';
if  TobFactaff <> Nil then
   BEGIN  TobFactaff.Free;  TobFactaff := Nil; END;
stSQL := 'SELECT '+chp+ ' From FACTAFF where AFA_AFFAIRE="' +  TObAffModele.getvalue('AFF_AFFAIRE') + '" AND AFA_TYPECHE="NOR"';
Q:= OpenSQL(stSQL,True);
if Not Q.EOF then
   BEGIN
    TobFactaff := TOB.Create('echeance des missions',Nil,-1);
    TobFactaff.LoadDetailDB ('FACTAFF','','',Q,False,False);
   END;
Ferme(Q);
END;





procedure TOF_AFAUGMEN.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
Var Statut : string;
begin

Statut := GetControltext('AFF_AFFAIRE0');

Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;



procedure AGLAugmen_Mission( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else laTof := nil;
if (Latof is TOF_AFAUGMEN) then TOF_AFAUGMEN(LaTof).Augmen_Mission else exit;
end;

Procedure AFLanceFiche_Mul_AugmentAff(Range:string);
begin
AGLLanceFiche ('BTP','BTAUGMEN_MUL',Range,'','STATUT=INT');
end;


Procedure TOF_AFAUGMEN.ControleChamp(Champ : String;Valeur : String);
Begin

  if Champ = 'STATUT' then
     Begin
     if Valeur = 'APP' then
        Begin
        Statut := 'W';
        Ecran.caption := 'Augmentation des Appels';
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appels');
        SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Appels');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appels');
        SetControlProperty('BEFFACEAFF2', 'Hint', 'Effacer Appels');
        SetControlText('TGPAFFAIRE', 'Code Appel');
        end
    Else if valeur = 'INT' then
        Begin
        Statut := 'I';
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
        SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Contrat');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Contrat');
        SetControlProperty('BEFFACEAFF2', 'Hint', 'Effacer Contrat');
        SetControlText('TGPAFFAIRE', 'Code Contrat');
    end
    Else if valeur = 'AFF' then
        Begin
      if assigned(GetControl('AFFAIRE0')) then SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("","A")');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
        SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Chantier');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Chantier');
        SetControlProperty('BEFFACEAFF2', 'Hint', 'Effacer Chantier');
        SetControlText('TGPAFFAIRE', 'Code Chantier');
        end
    else if Valeur = 'GRP' then
        Begin
      if assigned(GetControl('AFFAIRE0')) then SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
    end
    Else if valeur = 'PRO' then
    Begin
        Statut := 'P';
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
        SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Appel d''offre');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel d''offre');
        SetControlProperty('BEFFACEAFF2', 'Hint', 'Effacer Appel d''offre');
        SetControlText('TGPAFFAIRE', 'Code Appel d''offre');
        end
     else
        Begin
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel Affaire');
        SetControlProperty('BSELECTAFF2', 'Hint', 'Recherche Affaires');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Affaires');
        SetControlProperty('BEFFACEAFF2', 'Hint', 'Effacer Affaires');
        SetControlText('TGPAFFAIRE', 'Code Affaire');
        end;
    //
    if Statut <> '' then
    begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then
        SetControlText('AFF_AFFAIRE0', Statut)
      else if assigned(GetControl('AFFAIRE0')) then
        SetControlText('AFFAIRE0', Statut);
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="' + Statut + '"');
     end;
  end;

end;

FUNCTION TOF_AFAUGMEN.ChargeLesAffaires (Mul : TFMul; ChampCpt,TableCpt : string; TobEnregs : TOB; Parle : Boolean):T_ErrTraiteEnreg;
Var stWhere : string;
    i :integer;
begin
Result := tteOK;

if Mul = Nil then Exit;
if (TobEnregs = Nil) then Exit;

if Not(TestEnregAtraiter(Mul,Parle)) then
   begin Result := tteAucunEnreg; Exit; end;
{$IFDEF EAGLCLIENT}
if Mul.bSelectAll.Down then
   if not Mul.Fetchlestous then
     begin
     Mul.bSelectAllClick(Nil);
     Mul.bSelectAll.Down := False;
     exit;
     end;
{$ENDIF}

if Mul.FListe.AllSelected then
   begin
//   stWhere := ' ' + RecupWhereCritere(Mul.Pages);
   if Not (VerifNbrPourBlocage ( ChampCpt,TableCpt, stWhere , Parle, 200000 )) then
         begin Result := tteTropEnreg; Exit; end;
{$IFDEF EAGLCLIENT}
   TOBEnregs.LoadDetailDB('LISTE ENREG', '', '', Mul.Q.TQ,false)
{$ELSE}
   try
   THDBGrid(TFMUL(Mul).FListe).datasource.dataset.disablecontrols;
   if not Mul.Q.BOF then Mul.Q.First; // PL : Obligé de remettre au début car le loaddetaildb teste la position du curseur pas EOF
   TOBEnregs.LoadDetailDB('LISTE ENREG', '', '', Mul.Q,false);
   finally
   THDBGrid(TFMUL(Mul).FListe).datasource.dataset.enablecontrols;
   end;
{$ENDIF}
   end
else
   begin
   for i := 0 to Mul.FListe.NbSelected-1 do
      begin
      Mul.FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
      Mul.Q.TQ.Seek(Mul.FListe.Row-1) ;
      TOB.Create('Liste Enreg', TOBEnregs, -1).SelectDB('',Mul.Q.TQ, true);
{$ELSE}
      TOB.Create('Liste Enreg', TOBEnregs, -1).SelectDB('', Mul.Q, true);
{$ENDIF}
      end;
    end;
//******* fin de traitement déselection *******
// PL le 10/02/03 : Repositionne les flags de modification à False pour que les putvalue éventuels les positionnent à true
// et que si un update suit, ne soient updatés que les champs réellement modifiés par les put !  enfin,... j'me comprends
TOBEnregs.SetAllModifie (false);

// Gestion de la déselection des lignes et des boutons allSelect...
FinTraiteMul (Mul);
end;

Procedure TOF_AFAUGMEN.FinTraiteMul (Mul : TFMul);
begin
if Mul.FListe.AllSelected then Mul.FListe.AllSelected:=False
                          else Mul.FListe.ClearSelected;
Mul.bSelectAll.Down := False ;
end;

FUNCTION TOF_AFAUGMEN.TestEnregATraiter (Mul : TFMul; Parle : Boolean) : Boolean;
begin
result := True;
//Test si des enreg sont sélectionnés
if (Mul.FListe.NbSelected=0) and (not Mul.FListe.AllSelected) then
   begin
   if Parle then PGIBoxAf('Aucun élément sélectionné',TitreHalley);
   Result := False;
   end;
end;

FUNCTION TOF_AFAUGMEN.VerifNbrPourBlocage ( ChampCpt,TableCpt, stWhere : string;  Parle : Boolean ; nbEnregMax : integer) : Boolean;
Var   QQ : TQuery;
			SQl : string;
begin
Result := True;
if (ChampCpt = '') or (TableCpt = '') then Exit;
	Sql := 'SELECT count(' + ChampCpt + ') as NbrEnreg FROM AFAFFTIERS ' + stWhere;
QQ := OpenSQL (Sql, True);
if (Not QQ.EOF) and (QQ.FindField ('NbrEnreg').AsInteger > nbEnregMax) then
   begin
   Result := False;
   if Parle then PGIBoxAf(TexteMessage[2],TitreHalley);
   end;
Ferme(QQ);
end;


procedure TOF_AFAUGMEN.MoisDebChange(Sender: Tobject);
begin
  if MOISDEB.Value = '' then THEdit(GetCOntrol('XX_WHERE')).Text := ''
  else if MOISDEB.Value = '01' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=1)'
  else if MOISDEB.Value = '02' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=2)'
  else if MOISDEB.Value = '03' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=3)'
  else if MOISDEB.Value = '04' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=4)'
  else if MOISDEB.Value = '05' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=5)'
  else if MOISDEB.Value = '06' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=6)'
  else if MOISDEB.Value = '07' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=7)'
  else if MOISDEB.Value = '08' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=8)'
  else if MOISDEB.Value = '09' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=9)'
  else if MOISDEB.Value = '10' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=10)'
  else if MOISDEB.Value = '11' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=11)'
  else if MOISDEB.Value = '12' then THEdit(GetCOntrol('XX_WHERE')).Text := ' AND (MONTH(AFF_DATEDEBUT)=12)';
 THEdit(GetCOntrol('XX_WHERE')).Text:= THEdit(GetCOntrol('XX_WHERE')).Text + stNonresilie;
end;

Initialization
registerclasses ( [ TOF_AFAUGMEN ] ) ;
RegisterAglProc( 'Augmen_mission',True,0,AGLAugmen_Mission);
end.
