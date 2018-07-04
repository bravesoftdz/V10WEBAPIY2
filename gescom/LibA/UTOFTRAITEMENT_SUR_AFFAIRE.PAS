{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : TRAITEMENT_SUR_AFFAIRE ()
Mots clefs ... : TOF;TRAITEMENT_SUR_AFFAIRE
*****************************************************************}
Unit UTOFTRAITEMENT_SUR_AFFAIRE ;

Interface

Uses
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     FE_Main, Mul,dbTables, db,HDB,
{$ENDIF}
    StdCtrls, Controls, Classes,  forms, sysutils,
     ComCtrls,UtilPgi, UtilGc,
     HCtrls, HEnt1, Windows,HMsgBox,M3fp, HStatus, TiersUtil,
     UTOF,Utob,DicoAf,uTofAfBaseCodeAffaire,Traducaffaire,EntGC,Factutil,FactTOB,
     Ent1,AffaireUtil,AffEcheanceUtil,AffaireDuplic,AffaireModifPiece,
     UtilMulTrt,UtofAfTraducChampLibre,paramsoc;

Type
  TOF_AlignAff = Class (TOF_AFBASECODEAFFAIRE)
     procedure Align_Aff;
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override  ;
     Procedure OnArgument(stArgument : String ); override ;
  end ;

// Triatement épuration / génération écheances sur affaires
Type
  TOF_EpurEchAff = Class (TOF_AFBASECODEAFFAIRE)
     procedure Epur_EchAff;
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override  ;
     Procedure OnArgument(stArgument : String );  override ;
     Procedure OnUpdate;                          override ;
  private
     ZDateEpur: Thedit;
     DateDebGener, DateFinGener : TDateTime;
     bGener : Boolean;
  end ;

       // pour écran sélection alignement sur client
Type
     TOF_AlignSel = Class (TOF_AFTRADUCCHAMPLIBRE)
       public
       procedure OnArgument(stArgument : String ) ; override ;
     end;
Procedure AFLanceFiche_ALign_AffaireSurClt;
Procedure AFLanceFiche_EpurEchFact(Argument:string);
Function AFLanceFiche_Selec_AffSurClt:variant;


Implementation
//******************************************************************************
//********************* Gestion Alignement client ******************************
//******************************************************************************
// fct pour alignements clients sur affaire
procedure TOF_ALIGNAFF.OnArgument(stArgument : String );
begin
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}

// Suppression des champs mois de cloture en Gestion d'Affaire
if Not (ctxScot in V_PGI.PGIContexte) then
   begin       //mcd 24/02/03 ... mis bon nom de champ !!!
   SetControlVisible('T_MOISCLOTURE',False); SetControlVisible('T_MOISCLOTURE_',False);
   SetControlVisible('TT_MOISCLOTURE',False); SetControlVisible('TT_MOISCLOTURE_',False);
   end ;
inherited;
end;

 procedure TOF_ALIGNAFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));     Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));   Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));   Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));   Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_')); Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_')); Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));     Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure TOF_ALIGNAFF.Align_Aff;
var  F : TFMul ;
     QQ,QQ1:TQuery;
       TobAff,TObPiece : TOB;
       titre,Select,Stmp,maj, Sql:string;
       wi,piece :integer;
       Recep : TStringList;
       CleDoc : R_CLEDOC;
       dDateDebItv, dDateFinItv : TDateTime;
       ret:T_ErrTraiteEnreg;
begin
F:=TFMul(Ecran);
TobAff := Tob.Create('les affaires',NIL,-1);
// PA le 29/08/2001 - Fonction de traitement des enreg du mul externalisée
ret:=TraiteEnregMulListe (F, 'AFF_AFFAIRE','AFAFFTIERS', TobAff, True);// gm 22/08/02
If ret <> tteOK then exit ;  // mcd 15/06/02
// on sélectionne les champs à mettre à jour
Select := AFLanceFiche_Selec_AffSurClt ;
if (Select='') then
  begin
   Tobaff.free;

   exit;
   end;
sTmp := StringReplace(Select, ';', chr(VK_RETURN), [rfReplaceAll]);
Recep := TStringList.Create;
Recep.Text := sTmp;
InitMove(Tobaff.Detail.count*2,' ');
wi := 0;
while wi < TobAff.Detail.count  do
    Begin
    MoveCur(False);
    if (BlocageAffaire( 'MAF', TobAff.Detail[wi].GetValue('AFF_AFFAIRE'), V_PGI.groupe, V_PGI.DateEntree, False, False, false, dDateDebItv, dDateFinItv, nil)<>tbaAucun)
         then  begin
         PGIInfoAf(Format('Affaire %s %s non modifiable',[TobAff.Detail[wi].getValue('AFF_TIERS'),TobAff.Detail[wi].getValue('AFF_AFFAIRE')]),Titre);
         end
     else begin
       Piece := SelectPieceAffaire(TobAff.Detail[wi].GetValue('AFF_AFFAIRE'), TobAff.Detail[wi].GetValue('AFF_STATUTAFFAIRE'),Cledoc);
       if (Piece <> 1) then BEGIN PGIInfoAf('Nbre de pièces associées à la mission incorrecte',Titre);Exit; END;
          // ne récupère qu'un enrgt pour une affaire. on laisse
        QQ:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True) ;
        if Not (QQ.EOF) then
           BEGIN
           TOBPiece:=TOB.Create('PIECE',Nil,-1) ;
           TOBPiece.SelectDB('',QQ);
           if (recep.Values['RIB']='1') then begin
                // idem ne récupère qu'un enrgt pour la piece on laisse
              QQ1:=OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="'+TiersAuxiliaire(TobAff.Detail[wi].getValue('AFF_TIERS'),False)
                                   +'" AND R_PRINCIPAL="X"',True) ;
              if Not QQ1.EOF then TObpiece.putvalue('GP_RIB',EncodeRIB(QQ1.FindField('R_ETABBQ').AsString,QQ1.FindField('R_GUICHET').AsString,
                                                QQ1.FindField('R_NUMEROCOMPTE').AsString,QQ1.FindField('R_CLERIB').AsString,
                                                QQ1.FindField('R_DOMICILIATION').AsString)) ;
              Ferme(QQ1) ;
              end;
           maj :=TObpiece.Getvalue('GP_MAJLIBRETIERS');
           if (maj ='') then maj:='XXXXXXXXXXX';
           if (recep.Values['MODE']='1') then TObpiece.putvalue('GP_MODEREGLE',TobAff.Detail[wi].getValue('T_MODEREGLE'));
       // mcd 15/03/02 supprimer car il faut recalculer les lignes ...   if (recep.Values['REGIME']='1') then TObpiece.putvalue('GP_REGIMETAXE',TobAff.Detail[wi].getValue('T_REGIMETVA'));
           if (recep.Values['R1']='1') then TObpiece.putvalue('GP_TIERSSAL1',TobAff.Detail[wi].getValue('YTC_RESSOURCE1'));
           if (recep.Values['R2']='1') then TObpiece.putvalue('GP_TIERSSAL2',TobAff.Detail[wi].getValue('YTC_RESSOURCE2'));
           if (recep.Values['R3']='1') then TObpiece.putvalue('GP_TIERSSAL3',TobAff.Detail[wi].getValue('YTC_RESSOURCE3'));
           if ((recep.Values['T1']='1') and (Copy(Maj,1,1)<>'-')) then TObpiece.putvalue('GP_LIBRETIERS1',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS1'));
           if ((recep.Values['T2']='1') and (Copy(Maj,2,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS2',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS2'));
           if ((recep.Values['T3']='1') and (Copy(Maj,3,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS3',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS3'));
           if ((recep.Values['T4']='1') and (Copy(Maj,4,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS4',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS4'));
           if ((recep.Values['T5']='1') and (Copy(Maj,5,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS5',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS5'));
           if ((recep.Values['T6']='1') and (Copy(Maj,6,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS6',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS6'));
           if ((recep.Values['T7']='1') and (Copy(Maj,7,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS7',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS7'));
           if ((recep.Values['T8']='1') and (Copy(Maj,8,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS8',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS8'));
           if ((recep.Values['T9']='1') and (Copy(Maj,9,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERS9',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERS9'));
           if ((recep.Values['TA']='1') and (Copy(Maj,10,1)<>'-'))then TObpiece.putvalue('GP_LIBRETIERSA',TobAff.Detail[wi].getValue('YTC_TABLELIBRETIERSA'));
           if (recep.Values['TAR']='1') then TObpiece.putvalue('GP_TARIFTIERS',TobAff.Detail[wi].getValue('T_TARIFTIERS'));
           if (recep.Values['COMMERCIAL']='1')  then TObpiece.putvalue('GP_REPRESENTANT',TobAff.Detail[wi].getValue('T_REPRESENTANT'));

           TOBPiece.PutValue('GP_DATEMODIF', V_PGI.DateEntree);
           Try
            TobPiece.UpdateDb(False,False);
            if  ((recep.Values['TAR']='1') or (recep.Values['COMMERCIAL']='1')) then
              begin // mcd 20/01/03 il faut reporter ces 2 valeurs dans les lignes
              Sql := 'UPDATE LIGNE SET ';
              if  (recep.Values['TAR']='1')then Sql := Sql +' GL_TARIFTIERS="'+ TobAff.Detail[wi].getValue('T_TARIFTIERS')+'" ';
              if ((recep.Values['TAR']='1') and (recep.Values['COMMERCIAL']='1')) then Sql := Sql + ' , ';
              if  (recep.Values['COMMERCIAL']='1')then Sql := Sql +' GL_REPRESENTANT="'+ TobAff.Detail[wi].getValue('T_REPRESENTANT')+'" ';
              Sql := SQL + ' WHERE '+ WherePiece(CleDoc,ttdLigne,False);
              ExecuteSql (Sql);
              end;
           finally
            end;
           TobPiece.free;
           end
        else  PGIInfoAf(Format('Affaire %s non trouvée',[TobAff.Detail[wi].getValue('AFF_AFFAIRE')]),Titre);
       ferme(QQ);
       end;
    inc(wi);
    End;  // fin while  ,boucle sur la sélection

  Recep.Free;
  Tobaff.free;
  FiniMove;
end ;




//******************************************************************************
//**************** Gestion Epuration / génération des écheances client *********
//******************************************************************************

/// fct pour Epuration des échéances affaire facturées
procedure TOF_EPURECHAFF.OnArgument(stArgument : String );
begin
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
ZDateEpur:=THEdit(GetControl('ZDATEEPUR')) ;
// distinction Epuration / Génération  ( attention si plus d'arguments passés ...)
if stArgument = 'GENERATION' then bGener := true else bGener := False;
// impact en, affichage du mode épuration ou géneration
if bGener then
	SetControlProperty('BOuvrir','Hint', 'Génération des échéances')
else
	SetControlProperty('BOuvrir','Hint', 'Epuration des échéances');

SetControlVisible('ZDATEFIN',bGener);
SetControlVisible('ZDATEEPUR',not(bGener));
if bGener then
   begin
   SetControlText ('TZDATE', 'Génération jusqu''au ');
   Ecran.Caption := 'Génération des échéances d''affaires';
   UpdateCaption(Ecran);
//  SetControlText ('ZDATEDEB',DateTostr(Date));    //GM 22/08/02
//   SetControlText ('ZDATEFIN',DateToStr(PlusDate(Date,1,'A')));    //GM 22/08/02
 		SetControlText ('ZDATEFIN',DateToStr(GetParamSoc('SO_AFDATEFINGENER')));
   end;
// Suppression des champs mois de cloture en Gestion d'Affaire
if Not (ctxScot in V_PGI.PGIContexte) then
   begin
   SetControlVisible('T_MOISCLOTURE',False); SetControlVisible('T_MOISCLOTURE_',False);
   SetControlVisible('TT_MOISCLOTURE',False); SetControlVisible('TT_MOISCLOTURE_',False);
   end;
inherited;
end;

 procedure TOF_EPURECHAFF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));      Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));    Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));    Aff4:=THEdit(GetControl('AFF_AVENANT'));
Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));    Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));  Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));  Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));      Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;

procedure TOF_EPURECHAFF.OnUpdate ;
begin
  Inherited ;
  SetControlText('XX_WHERE',''); // gm le 20/10/03 pour ne pas que le mul envoie toutes les lignes
end ;

procedure TOF_EPURECHAFF.Epur_EchAff;
var  F : TFMul ;
     TobAff : TOB;
     titre,St:string;
     wi :integer;
     dDateDebItv, dDateFinItv : TDateTime;
     EdDate : THEdit;
     deb,fin,sttexte : string;
     ret:T_ErrTraiteEnreg;
begin

deb := FormatDateTime('dd/mm/yyyy ttttt',NowH);
F:=TFMul(Ecran);

{    //GM 22/08/02
if bGener then
   begin
   EdDate := THEdit(GetControl('ZDATEDEB'));
   if EdDate <> Nil then DateDebGener := StrToDate(EdDate.Text);
   EdDate := THEdit(GetControl('ZDATEFIN'));
   if EdDate <> Nil then DateFinGener := StrToDate(EdDate.Text);
   if (DateDebGener > DateFinGener)   or (DateFinGener = idate2099) or (DateDebGener=DateFinGener) then
      begin PGIBoxAf ('Dates de génération incorrectes',titre); Exit; End;
   end;
   }
if bGener then
   begin
   EdDate := THEdit(GetControl('ZDATEFIN'));
   if EdDate <> Nil then DateFinGener := StrToDate(EdDate.Text);
   if   DateFinGener = idate2099   then
      begin PGIBoxAf ('Dates de génération incorrectes',titre); Exit; End;
   end;

//if bGener then St := Format('Confirmez vous la génération des échéances du %s au %s',[DateTostr(DateDebGener),DateTostr(DateFinGener)])    //gm 22/08/02
if bGener then St := Format('Confirmez vous la génération des échéances jusqu''au %s',[DateTostr(DateFinGener)])
          else St := Format('Confirmez vous l''épuration des échéances facturées antérieures au %s',[ZdateEpur.text]);
If (PGIAskAf(st,titre)<> mrYes) then exit;
   // Ok pour faire le traitement
TobAff := Tob.Create('les affaires',NIL,-1);

// PA le 29/08/2001 - Fonction de traitement des enreg du mul externalisée
ret := TraiteEnregMulListe (F, 'AFF_AFFAIRE','AFAFFTIERS', TobAff, True);    //GM 22/08/02
If ret <> tteOK then exit ;
deb := FormatDateTime('dd/mm/yyyy ttttt',NowH);
InitMove(Tobaff.Detail.count,' ');
wi := 0;
while wi < TobAff.Detail.count  do
    Begin
    MoveCur(False);
    if ( (not VH_GC.GCIfDefCEGID) and (BlocageAffaire( 'MAF', TobAff.Detail[wi].GetValue('AFF_AFFAIRE'), V_PGI.groupe, V_PGI.DateEntree, False, False, false, dDateDebItv, dDateFinItv, nil)<>tbaAucun) ) or
       ( VH_GC.GCIfDefCEGID)
         then  begin
         PGIInfoAf(Format('Affaire %s %s non modifiable',[TobAff.Detail[wi].getValue('AFF_TIERS'),TobAff.Detail[wi].getValue('AFF_AFFAIRE')]),Titre);
         end
     else
         begin
         if bGener then  // **** Génération des écheances ****
            begin
            DateDebGener:= idate1900; //GM 22/08/02 afin de ne rien oublier
            GenerEcheancesAffaire (TobAff.Detail[wi].GetValue('AFF_AFFAIRE'),DateDebGener,DateFinGener);
            end
         else   // **** Epuration des écheances ****
            begin
            SupEcheancesAffaire (TobAff.Detail[wi].getValue('AFF_AFFAIRE'),True,StrToDate(ZdateEpur.text));
            AjusteDateGenerSurEch (TobAff.Detail[wi].GetValue('AFF_AFFAIRE'),True,True,StrToDate(ZdateEpur.text),idate1900);
            end;
         end;
    inc(wi);
    End;  // fin while  ,boucle sur la sélection

  TobAff.free;
  FiniMove;
	fin := FormatDateTime('dd/mm/yyyy ttttt',NowH);
  StTexte := Format('Traitement Ok de  %s à %s',[deb,fin]);
  if bGener then
	  PGIInfoAf(StTexte,'Génération des échéances')
  else
    PGIInfoAf(StTexte,'Epuration des échéances');
end;


procedure TOF_ALIGNSel.OnArgument(stArgument : String );
begin
// Suppression des champs mois de cloture en Gestion d'Affaire
if  (ctxScot in V_PGI.PGIContexte) then
   begin
   SetControlVisible('COMMERCIAL',False);
   end ;
if ((Ecran.FindComponent ('TYTC_TABLELIBRETIERS1'))<>nil) then
    begin  // mcd 25/07/02 les nom de la fiche ne sont pas std
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '_');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    end;
inherited;
end;

// ************ Appel depuis le script ************************
procedure AGLAlign_aff( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else latof := nil;
if (Latof is TOF_ALIGNAFF) then TOF_AlignAff(LaTof).Align_Aff else exit;
end;


procedure AGLEpurEchAff( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else latof := nil;
if (Latof is TOF_EPURECHAFF) then TOF_EPURECHAFF(LaTof).Epur_EchAff else exit;
end;

Procedure AFLanceFiche_ALign_AffaireSurClt;
begin
AGLLanceFiche('AFF','AFALIGNAFF','','','');
end;
Procedure AFLanceFiche_EpurEchFact(Argument:string);
begin
AGLLanceFiche('AFF','AFEPUR_ECH_AFF','','',Argument);
end;
Function AFLanceFiche_Selec_AffSurClt:variant;
begin
result:=AGLLanceFiche('AFF','AFALIGNSEL','','','');
end;


Initialization
RegisterClasses ( [TOF_AlignAff,TOF_EPurEchAff,TOF_AlignSel] ) ;
RegisterAglProc( 'Align_Aff',True,0,AGLAlign_Aff);
RegisterAglProc( 'Epur_EchAff',True,0,AGLEpurEchAff);
end.
