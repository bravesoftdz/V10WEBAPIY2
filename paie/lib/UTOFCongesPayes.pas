{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 28/08/2001
Modifié le ... :   /  /
Description .. : Gestion du multicritère des CP
Mots clefs ... : PAIE;CP
*****************************************************************
PT- 1 : 28/08/2001 : SB 547: Colle zero devant code salarie
PT- 2 : 21/09/2001 : SB 547: Moulinette de maj des CP
                          Nouvelles procedures
PT- 3 04/10/2001 SB 562 Modification du chemin de sauvegarde
PT- 4 09/10/2001 SB 562 Les mvts Ajp ne doivent pas être pris en
                        compte ds les consommés
PT- 5 26/09/2001 SB 562 Reconception du module absence
PT- 6 22/11/2001 SB 563 Pour ORACLE, Le champ PCN_CODETAPE doit être renseigné
PT- 7 26/11/2001 SB 563 Déplacement de la moulinette de maj des Cp déplacé ds un autre module
PT- 8 01/03/2002 SB 571 Appel de la Moulinette des maj des bases des mvts d'acquis congés payés
PT- 9 21/06/2002 PH 585 traitement en lookupapproximatif pour lookup salarié sinon erreur SQL
PT10-1 07/04/2004 SB V_50 FQ 11160 Ajout critère Congés pris
PT10-2                             Refonte design tableau récapitualtif
PT10-3                             Ajout Bouton Detail congés payés
PT10-4                             Raffraîchissement du caption de la fiche et non de l'étiquette etab
PT- 11 31/08/2004 JL V_50 FQ 11507 proposer par défaut : Mouvements <Tous> au lieu Congés pris
PT12              SB V_60 Détails CP, utilisation de la zone salarié de la fiche
PT13   20/06/2007 FC V_72 FQ 14196 Gestion du salarié confidentiel
PT14  29/06/2007  MF V_72 FQ 14470 Gestion CP par salarié : gestion des confidentiels
PT15  20/12/2007 FC V810 FQ 14996 Concepts accessibilité depuis la fiche salarié
PT16  03/08/2008 PH V810 FQ 15546 Positionnement sur le bon renregistrement en CWAS
}
unit UTOFCongesPayes;


interface
uses  Controls,Classes,sysutils,ComCtrls,
{$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,Fe_Main,db,
{$ELSE}
       eMul,MaineAgl,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,HTB97,ULibEditionPaie;

Type
     TOF_CongesPayes = Class (TOF)
       procedure OnArgument(stArgument: String); override;
       procedure OnLoad ;                        override ;
       procedure OnClose;                        override ;
       procedure BChercheClick(Sender: TObject);
       procedure EtabChange(Sender: TObject);
       procedure Calculcumul;
       procedure RemplaceLibelle;
       procedure RechercheLibelle;
       procedure RechercheOrdreSuivant;
       procedure GrilleDblClick(Sender: TObject);
       procedure BinsertClick(Sender: TObject);
       procedure ExitEdit(Sender: TObject);
       procedure LanceVerifCP(Sender: TObject); { PT- 2}  //PT- 7
       procedure LanceFicheDetail ( Sender : TObject );    { PT10-3 }
     private
       salarie,Etablissement,Origine,libOrdre,FicPrec,Action       : string;  //PT15
       Tob_Congespayes  : TOB;
       OrdreSuivant,PeriodeEnCours  : integer;
       DateCloture   : TdateTime;
// PT- 9 21/06/2002 PH 585 traitement en lookupapproximatif pour lookup salarié sinon erreur SQL
       PgLookUpAppro : Boolean;
       END ;

implementation
uses PgCongesPayes,UTOFPGUtilitaireCP,PGoutils2,EntPaie,P5Def;

procedure TOF_CongesPayes.OnArgument(stArgument: String);
var
Codetab  :  THValcombobox;
Defaut : ThEdit;
Arg : string;
Q,Q1 : TQuery;
T:TTabSheet;
Binsert,Bcherche : ttoolbarbutton97;
Zone : TControl;
num : integer;
Champ:string;
begin
Inherited ;
Etablissement:='';
// PT- 9 21/06/2002 PH 585 traitement en lookupapproximatif pour lookup salarié sinon erreur SQL
  PgLookUpAppro := V_PGI.LookUpLocate;
  if NOT V_PGI.LookUpLocate then V_PGI.LookUpLocate := TRUE;

  FicPrec := '';
  Arg := stArgument;
  salarie:=Trim(ReadTokenPipe(Arg,';')) ;
  Origine:=Trim(ReadTokenPipe(Arg,';')) ;
  etablissement := Trim(ReadTokenPipe(Arg,';')) ;
  libordre := Trim(ReadTokenPipe(Arg,';')) ;
  //DEB PT15
  if Arg <> '' then
    FicPrec := Trim(ReadTokenPipe(Arg,';')) ;
//  FicPrec := Arg;
  Action := '';
  if Arg <> '' then
  begin
    if pos('=', Arg) <> 0 then
    begin
      Champ := ReadTokenPipe(Arg, '=');
      if (Champ = 'ACTION') then
        Action := Arg;
    end;
  end;
  //FIN PT15

  if Salarie='' then SetcontrolText('PCN_TYPECONGE','');   { PT10-1 }  //PT- 11
  SetcontrolText('PGMVTPRIS','PAR');

  BCherche := TToolbarbutton97(getcontrol('BDETAIL'));
  if BCherche <> nil then  BCherche.onclick := LanceFicheDetail;   { PT10-3 }

  TFMul(Ecran).FListe.OnDblClick := GrilleDblClick; {PT- 5} //PT15

  SetControlVisible('BVERIF',V_PGI.Debug ); {DEB2 PT- 2    PT- 7 V_PGI.Debug   }
  Binsert := ttoolbarbutton97(getcontrol('BVERIF')); //PT- 7 mise en commentaire
  if Binsert <> nil then Binsert.OnClick := LanceVerifCP; {FIN2 PT- 2}

  Binsert := ttoolbarbutton97(getcontrol('Binsert'));
  if Binsert <> nil then Binsert.OnClick := BinsertClick;

  Defaut:=ThEdit(getcontrol('PCN_SALARIE'));    {PT- 1}
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;

  SetControlText('PCN_PERIODECP','0');
  SetControlText('PCN_SALARIE',Salarie);
  if Salarie <> '' then
     begin
     Q:=opensql('SELECT PSA_ETABLISSEMENT, PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE = "'+ SALARIE+'"',TRUE);     if not Q.eof then
     If not Q.Eof then //PORTAGECWAS
       Begin
       Ecran.Caption := 'Salarié ' + SALARIE + ' ' + Q.FindField('PSA_LIBELLE').Value;
       Etablissement := Q.FindField('PSA_ETABLISSEMENT').Value;
       End;    
     ferme(Q);
     end;

  Zone:=ThValComboBox(getcontrol('PCN_ETABLISSEMENT'));
  InitialiseCombo(Zone);

  For Num := 1 to 4 do
    VisibiliteChampSalarie (IntToStr(Num),GetControl ('PCN_TRAVAILN'+IntToStr(Num)),GetControl ('TPCN_TRAVAILN'+IntToStr(Num)));
  VisibiliteStat (GetControl ('PCN_CODESTAT'),GetControl ('TPCN_CODESTAT')) ;

  EtabChange(nil);
  if etablissement<>'' then
    Begin
    Q1 := opensql('SELECT ETB_PERIODECP,ETB_DATECLOTURECPN FROM ETABCOMPL '+
    'WHERE ETB_ETABLISSEMENT = "'+etablissement + '"', TRUE);
    if Q1.eof then begin PeriodeEnCours:=0; Ferme(Q1); exit; end;
    PeriodeEnCours := Q1.findfield('ETB_PERIODECP').asVariant;
    DateCloture := Q1.findfield('ETB_DATECLOTURECPN').AsVariant;
    Ferme(Q1);
    end
    else exit;

  T := TTabSheet(getcontrol('PCOMPLEMENT'));
  if T <> nil then
     T.tabvisible := false;

  T := TTabSheet(getcontrol('PAVANCE'));
  if T <> nil then
     T.tabvisible := false;

  BCherche := TToolbarbutton97(getcontrol('BCHERCHE'));
  if BCherche <> nil then
     begin
     BCherche.onclick := BChercheClick;
     end;


  Codetab := THValComboBox(getcontrol('PCN_ETABLISSEMENT'));
  if Codetab <> nil then  Codetab.onchange := EtabChange;
  setcontrolVisible('PCN_SALARIE'          ,(FicPrec<>'BULL')and (FicPrec <> 'SAL'));
  setcontrolVisible('TPCN_SALARIE'         ,(FicPrec<>'BULL')and (FicPrec <> 'SAL'));
  setcontrolVisible('PCN_ETABLISSEMENT'    ,(FicPrec<>'BULL')and (FicPrec <> 'SAL'));
  setcontrolVisible('TPCN_ETABLISSEMENT'   ,(FicPrec<>'BULL')and (FicPrec <> 'SAL'));
  setcontrolVisible('PCN_LIBELLE'          ,(FicPrec<>'BULL')and (FicPrec <> 'SAL'));
  setcontrolVisible('TPCN_LIBELLE'         ,(FicPrec<>'BULL')and (FicPrec <> 'SAL'));

  //DEB PT15
  if Action='CONSULTATION' then
    SetControlEnabled('BInsert',False);
  //FIN PT15
end;

procedure TOF_CongesPayes.BinsertClick(Sender: TObject);
var
  St : String;
begin
  if Origine = 'C' then
    begin
    if GetControlText('PCN_SALARIE') = '' then
      begin
      PGIBox('Vous devez saisir un matricule salarié.','Création mouvement de congés');
      SetFocusControl('PCN_SALARIE');
      exit;
      end
    else
    begin
      Salarie:=GetControlText('PCN_SALARIE');
      //DEB PT13
      St := SQLConf('SALARIES');
      if St <> '' then
        St := ' AND ' + St;
      if not existesql('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE= "' + salarie + '"' + St) then
      begin
        PGIBox('Le matricule salarié n''existe pas.','Création mouvement de congés');
        SetFocusControl('PCN_SALARIE');
        exit;
      end;
    end;
      //FIN PT13
     // Ressource:=TFMUL(Ecran).Q.findfield('PCN_RESSOURCE').Asstring;
    end;
//   if V_PGI.Debug then
//     AglLanceFiche ('PAY','MVTCONGEPAYES',''  , '', salarie+';'+Ressource+';'+Origine+';'+ETABLISSEMENT +';ACTION=CREATION;'+libordre)
//   Else
     AglLanceFiche ('PAY','MVTCONGEPAYE',''  , '', salarie+';'+Origine+';'+ETABLISSEMENT +';ACTION=CREATION;'+libordre);

   TFMUL(Ecran).BChercheClick (Sender);
end;
procedure TOF_CongesPayes.EtabChange(Sender: TObject);
var
  CodSal   :  THEdit;
  CodEtab :  THValComboBox;
begin
  if (FicPrec = 'SAL') or (FicPrec='BULL') then exit;
  CodSal := THEdit(getcontrol('PCN_SALARIE'));
  if CodSal <> nil then
     CodSal.text := '';
  CodEtab :=  THValComboBox(getcontrol('PCN_ETABLISSEMENT'));
  if ((Codsal<>nil) and (CodEtab<>nil) and (CodEtab.value <>'')) then
     Codsal.plus := ' PSA_ETABLISSEMENT = "'+CodEtab.Value+'"';

end;
procedure TOF_CongesPayes.GrilleDblClick(Sender: TObject);
var
Ordre,typemvt : string;
begin
  //DEB PT15
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1); // PT16
{$ENDIF}
  Ordre := inttostr(TFMUL(Ecran).Q.findfield('PCN_ORDRE').Asinteger);
  Typemvt :=TFMUL(Ecran).Q.findfield('PCN_TYPEMVT').Asstring;
  if Origine = 'C' then
    Salarie := TFMUL(Ecran).Q.findfield('PCN_SALARIE').Asstring;
//    AglLanceFiche ('PAY','MVTCONGEPAYE',''  , typemvt+';'+salarie+';'+Ordre, salarie+';'+Origine+ ';'+ ETABLISSEMENT+';ACTION=MODIFICATION;'+libordre);

  if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
  begin
    if (Not JaiLeDroitTag(200067)) and (Action='CONSULTATION') then
      AglLanceFiche ('PAY','MVTCONGEPAYE',''  , 'CPA;'+Salarie+';'+Ordre, Salarie+';C;'+ETABLISSEMENT+';ACTION=CONSULTATION;MONOFICHE')
    else
      AglLanceFiche ('PAY','MVTCONGEPAYE',''  , 'CPA;'+Salarie+';'+Ordre, Salarie+';C;'+ETABLISSEMENT+';ACTION=MODIFICATION;MONOFICHE');
  end
  else
  begin
    if (Not JaiLeDroitTag(200067)) and (Action='CONSULTATION') then
      AglLanceFiche ('PAY','MVTCONGEPAYE',''  , 'CPA;'+Salarie+';'+Ordre, Salarie+';C;'+ETABLISSEMENT+';ACTION=CONSULTATION')
    else
      AglLanceFiche ('PAY','MVTCONGEPAYE',''  , 'CPA;'+Salarie+';'+Ordre, Salarie+';C;'+ETABLISSEMENT+';ACTION=MODIFICATION');
  end;

  TFMUL(Ecran).BChercheClick (Sender);
  //FIN PT15
end;
procedure TOF_CongesPayes.OnLoad ;
// d PT14
var
   StConf       : string;
// f PT14
begin

  RechercheLibelle;
// d PT14
  StConf := '';
  StConf := SqlConf ('ABSENCESALARIE');
  if (StConf <> '') then
    StConf := 'AND '+ StConf;
// f PT14

  if GetcontrolText('PGMVTPRIS')= 'PAR' then            { PT10-1 }
    SetcontrolText('XX_WHERE',' PCN_MVTDUPLIQUE<>"X" '+StConf)  //PT14
  Else
    if GetcontrolText('PGMVTPRIS')= 'ECL' then
    SetcontrolText('XX_WHERE',' PCN_CODERGRPT<>-1 '+StConf);  //PT14
  { Pour test
  if GetControlText('PCN_SALARIE')<>'' then
    Begin
    //Acquis
    EXECUTESQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="ACQ" '+
               'WHERE PCN_SALARIE="'+GetControlText('PCN_SALARIE')+'" '+
               'AND PCN_TYPEMVT="CPA" AND PCN_APAYES<>PCN_JOURS '+
               'AND PCN_TYPECONGE<>"PRI" AND PCN_TYPECONGE<>"AJP" AND PCN_TYPECONGE<>"AJU" AND PCN_TYPECONGE<>"CPA" ');
    //Acquis consommé
    EXECUTESQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="CON" '+
               'WHERE PCN_SALARIE="'+GetControlText('PCN_SALARIE')+'" '+
               'AND PCN_TYPEMVT="CPA" AND PCN_APAYES=PCN_JOURS');
    //Pris
    EXECUTESQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="PRI" '+
               'WHERE PCN_SALARIE="'+GetControlText('PCN_SALARIE')+'" '+
               'AND PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_CODETAPE="..." ');
    //Pris payés
    EXECUTESQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="PAY" '+
               'WHERE PCN_SALARIE="'+GetControlText('PCN_SALARIE')+'" '+
               'AND PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_CODETAPE<>"..." ');
    //Pris éclaté
    EXECUTESQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="ECL" '+
               'WHERE PCN_SALARIE="'+GetControlText('PCN_SALARIE')+'" '+
               'AND PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_CODERGRPT=-1 ');
    //Mouvement cloturé
    EXECUTESQL('UPDATE ABSENCESALARIE SET PCN_ETATPOSTPAIE="CLO" '+
               'WHERE PCN_SALARIE="'+GetControlText('PCN_SALARIE')+'" '+
               'AND PCN_TYPEMVT="CPA" AND PCN_PERIODECP>1');
    End; }

end;
procedure TOF_CongesPayes.RechercheLibelle ;
var
   Periode   :  String;
   DDebut,Dfin : TDateTime;
   Acquis,Pris, Restants,Base,MBase: double;
   Q1 : TQuery;
Begin
// affichage du libellé correspondant à la période sélectionnée
  Salarie:=GetControlText('PCN_SALARIE');
  etablissement:=GetControlText('PCN_ETABLISSEMENT');
  Periode:=GetControlText('PCN_PERIODECP');
  if etablissement <> '' then
    begin
    Q1 := OpenSql('SELECT ETB_DATECLOTURECPN FROM ETABCOMPL '+
         'WHERE ETB_ETABLISSEMENT = "'+etablissement + '"', TRUE);
    if not Q1.eof then
      DateCloture := Q1.findfield('ETB_DATECLOTURECPN').Asdatetime;
    ferme(Q1);
    end;
  if ((Periode = '')or (etablissement='')) then
       Begin
       Ecran.caption := 'Liste des congés payés : '; { PT10-4 }
       UpdateCaption(Ecran);
       End
   else
     begin
     RechercheDate(DDebut,Dfin,DateCloture,Periode);
     Ecran.caption := 'Liste des congés payés du '+ Datetostr(DDebut)+' au '+ Datetostr(DFin); { PT10-4 }
     UpdateCaption(Ecran);
     end;
// affichage des cumuls pour la période sélectionnée
  if Periode='' then
    begin
     SetControlText('LPRIS',    'Non significatif');
     SetControlText('LACQUIS',  'Non significatif');
     SetControlText('LRESTANTS','Non significatif');
     SetControlText('LBASE',    'Non significatif');
    end
  else
    begin
     //SB 28/05/2002 Ajout Variable Idate1900 suite modif function
     AffichelibelleAcqPri(Periode,Salarie,Idate1900,0,Pris,Acquis,Restants,Base,MBase,false,False);
     SetControlText('LPRIS',    Format('%2.2f',[Pris]));        { PT10-2 }
     SetControlText('LACQUIS',  Format('%2.2f',[Acquis]));      { PT10-2 }
     SetControlText('LRESTANTS',Format('%2.2f',[Restants]));    { PT10-2 }
     SetControlText('LBASE',    Format('%2.2f',[Base]));        { PT10-2 }
    end;
end;

procedure  TOF_CongesPayes.Calculcumul;
var
   T : Tob;
   Cum_jours,Cum_base,jours,base : double;
   sens :string;
begin
   // Calcul des cumuls
   Cum_jours:=0;Cum_base :=0;
   T := TOb_congespayes.FindFirst([''],[''], true);
   while  T <> nil
   do begin
        sens   := T.getvalue('PCN_SENSABS');
        jours  := T.getvalue('PCN_JOURS');
        base   :=T.getvalue('PCN_BASE');
        if sens = '+' then
          begin
          Cum_jours   := Cum_jours +jours;
          Cum_base    := Cum_base+ base;
          end
          else
          begin
          Cum_jours   := Cum_jours -jours;
          Cum_base    := Cum_base- base;
          end;
        T := TOb_congespayes.Findnext([''],[''], true);
   end;
   SetControlText('CUM_JOURS',floattostr(cum_jours));
   SetControlText('CUM_BASE',floattostr(cum_Base));
End;

procedure TOF_CongesPayes.RemplaceLibelle;
var
T : Tob;
begin

 T := TOb_congespayes.FindFirst([''],[''],false);
 while  T <> nil  do
 begin
   T.PutValue('PCN_TYPECONGE',Rechdom('PGMVTCONGES',T.getvalue('PCN_TYPECONGE'),false));
   T.UpdateDB(false);
   T := TOb_congespayes.FindNext([''],[''],false);
 end;
end;


// changement des critères de sélection des mvts
procedure TOF_CongesPayes.BChercheClick(Sender: TObject);
begin
TFMul(Ecran).BChercheClick(Sender);
Salarie:=GetControlText('PCN_SALARIE');
Recherchelibelle;
end;


procedure  TOF_CongesPayes.RechercheOrdreSuivant;
var
  i : integer;
  TT : Tob;
begin
    OrdreSuivant:=0;
    TT := Tob_Congespayes.FindFirst([''],[''],True);
    while TT <> nil do
       begin
       i := TT.GetValue('PCN_ORDRE');
       if i > OrdreSuivant then
            OrdreSuivant := i;
       TT := Tob_Congespayes.FindNext([''],[''],True);
       end;
    OrdreSuivant:=OrdreSuivant+1;
end;

procedure TOF_CongesPayes.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end; 

procedure TOF_CongesPayes.LanceVerifCP(Sender: TObject);
begin
VerificationCP(GetControlText('PCN_SALARIE'),GetControlText('PCN_ETABLISSEMENT'));
end;
// PT- 9 21/06/2002 PH 585 traitement en lookupapproximatif pour lookup salarié sinon erreur SQL
procedure TOF_CongesPayes.OnClose;
begin
  inherited;
  V_PGI.LookUpLocate :=  PgLookUpAppro ;

end;

{ DEB PT10-3 }
procedure TOF_CongesPayes.LanceFicheDetail(Sender: TObject);
var
  St : String;
begin
  IF GetControlText('PCN_SALARIE')='' then
  Begin
    PGIBox('Vous devez saisir un matricule salarié.','Détails des congés payés');
    SetFocusControl('PCN_SALARIE');
    exit;
  End
  else
  begin
    //DEB PT13
    St := SQLConf('SALARIES');
    if St <> '' then
      St := ' AND ' + St;
    if not existesql('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE= "' + GetControlText('PCN_SALARIE') + '"' + St) then
    begin
      PGIBox('Le matricule salarié n''existe pas.','Détails des congés payés');
      SetFocusControl('PCN_SALARIE');
      exit;
    end;
    //FIN PT13
  end;
  AglLanceFiche('PAY','CONGESDETAILS','','',GetControlText('PCN_SALARIE')); { PT12 }
end;
{ FIN PT10-3 }

Initialization
registerclasses([TOF_CongesPayes]) ;
end.



