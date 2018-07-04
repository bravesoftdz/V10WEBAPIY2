{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_CURSUS ()
Mots clefs ... : TOF;PGMUL_CURSUS
*****************************************************************
PT1 | 24/11/2003 | V_50  | JL | Ajout Q.TQ.Seek en CWAS
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2 | 18/09/2007 | V_800 | FL | Ne pas gérer le double-clic dans la grille si aucun cursus visible
PT3 | 28/09/2007 | V_7   | FL | Emanager / Report / Adaptation cursus + accès assistant
PT4 | 05/10/2007 | V_7   | FL | Emanager / Report / Affichage du nombre d'heures de chaque stage + visu détail sur double-clic + plus de programme
PT5 | 24/10/2007 | V_8   | FL | Emanager / Report / Gestion des favoris
PT6 | 15/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT7 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT8 | 17/04/2008 | V_804 | FL | Ajout de PGBundleCatalogue
PT9 | 17/04/2008 | V_804 | FL | Ne pas autoriser la modification d'éléments STD si pas droits multi
}
Unit UTofPGMul_Cursus ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBGrids,Mul,FE_Main,
{$ELSE}
     eMul,MaineAGL,
{$ENDIF}
     UTOB,forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry,HTB97,PGOutilsFormation ;

Type
  TOF_PGMUL_CURSUS = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Argument{$IFDEF EMANAGER},typeUtilisat{$ENDIF},MillesimeEC : String;
    procedure GrilleDblClick (Sender : TObject) ;
    procedure Nouveau(Sender : TObject);
  end ;

    TOF_PGMUL_CURSUSEFOR = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Argument{$IFDEF EMANAGER},TypeUtilisat{$ENDIF},MillesimeEC : String;
    procedure GrilleDblClick (Sender : TObject) ;
  end ;

  TOF_PGEDITCURSUS = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ; //PT6
  end;

  //PT3 - Début
  {$IFDEF EMANAGER}
  TOF_PGEM_MULCATCURSUS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Argument,MillesimeEC : String;
    procedure GrilleDblClick (Sender : TObject) ;
    procedure AjouterFavoris (Sender : TObject);
  End;

  TOF_PGEMCURSUS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Cursus, Rang : String;
    MillesimeEC : String;
    Procedure GrilleDblClick (Sender : TObject);    //PT3
    Procedure AjouterAuxFavoris (Sender : TObject); //PT5
    Procedure RetirerDesFavoris(Sender : Tobject);  //PT5
  End;
  {$ENDIF}
  //PT3 - Fin
Implementation                                  
Uses uTableFiltre,SaisieList,GalOutil;                                  

{TOF_PGMUL_CURSUS}

procedure TOF_PGMUL_CURSUS.Nouveau(Sender : TObject);
begin
        AglLanceFiche('PAY', 'MUL_CURSUSEFOR', '', '', '');
end;

procedure TOF_PGMUL_CURSUS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_CURSUS.OnLoad ;
begin
  Inherited ;
        {$IFNDEF EMANAGER}
        If GetControlText('LISTECURSUS') = 'PREV' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS IN (SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+GetControlText('MILLESIME')+'")')
        else If GetControlText('LISTECURSUS') = 'CAT' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS NOT IN (SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+GetControlText('MILLESIME')+'")')
        else SetControlText('XX_WHERE','PCU_INCLUSCAT="X"');

        SetControLText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PCU_PREDEFINI'),GetControlText('NODOSSIER'),'PCU')); //PT6

        {$ELSE}
        SetControlText('LISTECURSUS','PREV');
        SetControlText('MILLESIME',MillesimeEC);
        If TypeUtilisat = 'R' then
        begin
                If GetControlText('LISTECURSUS') = 'PREV' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'")')
                else If GetControlText('LISTECURSUS') = 'CAT' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS NOT IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'")')
                else SetControlText('XX_WHERE','PCU_INCLUSCAT="X"');
        end
        else
        begin
                If GetControlText('LISTECURSUS') = 'PREV' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))')
                else If GetControlText('LISTECURSUS') = 'CAT' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS NOT IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))')
                else SetControlText('XX_WHERE','PCU_INCLUSCAT="X"');
        end;
        {$ENDIF}
end ;

procedure TOF_PGMUL_CURSUS.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        BNew : TToolBarButton97;
begin
  Inherited ;
        Argument := ReadTokenPipe(S,';');
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;

       BNew := TToolBarButton97(Getcontrol('BInsert'));
       If BNew <> Nil then BNew.OnClick := Nouveau;
       SetControlText('LISTECURSUS','PREV');
       Ecran.Caption := 'Inscriptions par cursus';
       UpdateCaption(Ecran);
       MillesimeEC := RendMillesimePrevisionnel;
       SetControlText('MILLESIME',MillesimeEC);
       {$IFDEF EMANAGER}
       If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
       else TypeUtilisat := 'S';
       SetControlVisible('BINSERT',True);
       SetControlVisible('BParamListe',False);
       SetControlVisible('PCU_CURSUS_',False);
       SetControlVisible('TPCU_CURSUS_',False);
       SetControlCaption('TPU_CURSUS','Cursus');
       SetControlVisible('PAvance',False);
       SetControlProperty('PAvance','TabVisible',False);
       SetControlEnabled('MILLESIME',False);
       SetControlEnabled('LISTECURSUS',False);
       {$ENDIF}

       {$IFNDEF EMANAGER}
       //PT6 - Début
       If PGBundleCatalogue then //PT8
       begin
         If not PGDroitMultiForm then
         begin
           SetControlEnabled('NODOSSIER',False);
           SetControlText   ('NODOSSIER',V_PGI.NoDossier);
           SetControlEnabled('PCU_PREDEFINI',False);
           SetControlText   ('PCU_PREDEFINI','');
           SetControlProperty ('PCU_CURSUS',  'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True)); //PT7
           SetControlProperty ('PCU_CURSUS_', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True)); //PT7
         end
       	 Else If V_PGI.ModePCL='1' Then 
       	 Begin
       	 	SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT7
           SetControlProperty ('PCU_CURSUS',  'Plus', DossiersAInterroger('','','PCU',True,True)); //PT7
           SetControlProperty ('PCU_CURSUS_', 'Plus', DossiersAInterroger('','','PCU',True,True)); //PT7
       	End;
       end
       else
       begin
         SetControlVisible('PCU_PREDEFINI', False);
         SetControlVisible('TPCU_PREDEFINI',False);
         SetControlVisible('NODOSSIER',     False);
         SetControlVisible('TNODOSSIER',    False);
       end;
       //PT6 - Fin
       {$ENDIF}
end ;

procedure TOF_PGMUL_CURSUS.GrilleDblClick (Sender : TObject) ;
var St : String ;
    Q_Mul : THQuery ;
begin
        If GetControlText('MILLESIME') = '' then
        begin
                PGIBox('Vous devez choisir un millésime',Ecran.Caption);
                Exit;
        end;
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;  //PT1
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q')) ;

        If Q_Mul.FindField('PCU_CURSUS').AsString = '' Then Exit; //PT2

        St := Q_Mul.FindField('PCU_CURSUS').AsString+';'+GetControlText('MILLESIME');
        AGLLanceFiche('PAY','INSCCURSUSPREV','','','CURSUS;'+St+';ACTION=MODIFICATION');
        //TFMul(Ecran).BCherche.Click;
end ;

{TOF_PGMUL_CURSUSEFOR}
procedure TOF_PGMUL_CURSUSEFOR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_CURSUSEFOR.OnLoad ;
begin
  Inherited ;
        {$IFNDEF EMANAGER}
        If GetControlText('LISTECURSUS') = 'PREV' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS IN (SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+GetControlText('MILLESIME')+'")')
        else If GetControlText('LISTECURSUS') = 'CAT' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS NOT IN (SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+GetControlText('MILLESIME')+'")')
        else SetControlText('XX_WHERE','PCU_INCLUSCAT="X"');
        {$ELSE}
        SetControlText('LISTECURSUS','CAT');
        SetControlText('MILLESIME',MillesimeEC);
        If TypeUtilisat = 'R' then
        begin
                If GetControlText('LISTECURSUS') = 'PREV' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'")')
                else If GetControlText('LISTECURSUS') = 'CAT' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS NOT IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'")')
                else SetControlText('XX_WHERE','PCU_INCLUSCAT="X"');
        end
        else
        begin
                If GetControlText('LISTECURSUS') = 'PREV' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))')
                else If GetControlText('LISTECURSUS') = 'CAT' then SetControlText('XX_WHERE','PCU_INCLUSCAT="X" AND PCU_CURSUS NOT IN (SELECT DISTINCT PCC_CURSUS FROM CURSUSSTAGE LEFT JOIN INSCFORMATION ON PFI_MILLESIME=PCC_MILLESIME AND PCC_CURSUS=PFI_CURSUS'+
                ' WHERE PFI_MILLESIME="'+GetControlText('MILLESIME')+'" AND PFI_CODESTAGE="--CURSUS--" AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))')
                else SetControlText('XX_WHERE','PCU_INCLUSCAT="X"');
        end;
        {$ENDIF}
end ;

procedure TOF_PGMUL_CURSUSEFOR.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
begin
  Inherited ;
        Argument := ReadTokenPipe(S,';');
        {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        SetControlVisible('BINSERT',False);
       SetControlText('LISTECURSUS','CAT');
       MillesimeEC := RendMillesimePrevisionnel;
       SetControlText('MILLESIME',MillesimeEC);
       Ecran.Caption := 'Liste des cursus';
       UpdateCaption(Ecran);
       {$IFDEF EMANAGER}
       If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
       else TypeUtilisat := 'S';
       SetControlVisible('BParamListe',False);
       SetControlVisible('PCU_CURSUS_',False);
       SetControlVisible('TPCU_CURSUS_',False);
       SetControlCaption('TPU_CURSUS','Cursus');
       SetControlVisible('PAvance',False);
       SetControlProperty('PAvance','TabVisible',False);
       SetControlEnabled('LISTECURSUS',False);
       SetControlEnabled('MILLESIME',False);
       {$ENDIF}
end ;

procedure TOF_PGMUL_CURSUSEFOR.GrilleDblClick (Sender : TObject) ;
var St : String ;
    Q_Mul : THQuery ;
begin
        If GetControlText('MILLESIME') = '' then
        begin
                PGIBox('Vous devez choisir un millésime',Ecran.Caption);
                Exit;
        end;
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;  //PT1
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
        St := Q_Mul.FindField('PCU_CURSUS').AsString+';'+GetControlText('MILLESIME');
        AGLLanceFiche('PAY','INSCCURSUSPREV','','','CURSUS;'+St+';ACTION=MODIFICATION');
        //If Retour = 'RAFRAICHIR' then TFMul(Ecran).BCherche.Click;
end ;

{TOF_PGEDITCURSUS}
procedure TOF_PGEDITCURSUS.OnArgument (S : String ) ;
begin
  Inherited ;
	If PGBundleCatalogue then //PT8
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier); //PT6
			SetControlEnabled('PCU_PREDEFINI',False);
			SetControlText   ('PCU_PREDEFINI','');
			SetControlProperty ('PCU_CURSUS',  'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True)); //PT9
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT7
	end
	else
	begin
		SetControlVisible('PCU_PREDEFINI', False);
		SetControlVisible('TPCU_PREDEFINI',False);
		SetControlVisible('NODOSSIER',     False); //PT6
		SetControlVisible('TNODOSSIER',    False);
	end;
end;

procedure TOF_PGEDITCURSUS.OnLoad ;
begin
	SetControLText('XX_WHERE', DossiersAInterroger(GetControlText('PCU_PREDEFINI'),GetControlText('NODOSSIER'),'PCU')); //PT6
end;

//PT3 - Début
{ TOF_PGEM_MULCATCURSUS }
{$IFDEF EMANAGER}

procedure TOF_PGEM_MULCATCURSUS.OnArgument(S: String);
var Liste : THGrid ;
    Bt    : TToolbarButton97;
begin
  inherited;
       Argument := ReadTokenPipe(S,';');

       Liste := THGrid(GetControl('FListe'));
       If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;

       MillesimeEC := RendMillesimeEManager;

       If Argument = 'INSC' Then
       Begin
            Ecran.Caption := 'Inscriptions aux cursus';
            UpdateCaption(Ecran);
       End
       //PT5 - Début
       else
       begin
             Bt := TToolBarButton97(GetControl('BFAVORIS'));
             If Bt <> Nil then
             begin
                  Bt.OnClick := AjouterFavoris;
                  Bt.Visible := True;
             end;
             if Liste  <>  NIL then Liste.OnDblClick  :=  GrilleDblClick;
       end;

       If ExisteSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_MILLESIME="'+MillesimeEC+'" AND PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'"') then
          SetControlChecked('FAVORIS',True)
       Else
          SetControlChecked('FAVORIS',False);
       //PT5 - Fin

       SetControlProperty('CBSTAGE','Plus','AND PST_FORMATION3="'+CYCLE_EN_COURS_EM+'"');
end;

procedure TOF_PGEM_MULCATCURSUS.OnLoad;
var StWhere : String;
begin
  inherited;
        If GetControlText('CBSTAGE') <> '' Then
                StWhere := 'PCU_CURSUS IN (SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_CODESTAGE="'+GetControlText('CBSTAGE')+'")'
        Else
                StWhere := '';

        //PT5 - Début
        If GetCheckBoxState('FAVORIS') = CbChecked then
        begin
             If StWhere <> '' then StWhere := StWhere+ ' AND ';
             StWhere := StWhere + 'PCU_CURSUS IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
             SetControlEnabled('BFAVORIS', False);
        end
        Else
        begin
             If StWhere <> '' then StWhere := StWhere+ ' AND ';
             StWhere := StWhere + 'PCU_CURSUS NOT IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
             SetControlEnabled('BFAVORIS', True);
        end;
        //PT5 - Fin

        SetControlText('XX_WHERE', StWhere);
end;

procedure TOF_PGEM_MULCATCURSUS.GrilleDblClick(Sender: TObject);
var St : String;
    Q_Mul : THQuery;
begin
{$IFDEF EAGLCLIENT}
       TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;
{$ENDIF}       
       Q_Mul := THQuery(Ecran.FindComponent('Q')) ;

       If Argument = 'INSC' Then
       Begin
            St := Q_Mul.FindField('PCU_CURSUS').AsString + ';' + RendMillesimeEManager;
            AGLLanceFiche('PAY', 'EM_SAISIECURSUS', '', '', 'CURSUS;'+St);
       End
       Else If Argument = 'CAT' Then
       Begin
            St := Q_Mul.FindField('PCU_CURSUS').AsString + ';0';
            AGLLanceFiche('PAY', 'EM_CURSUS', '', St, 'ACTION=CONSULTATION;'+St+';'+Q_Mul.FindField('PCU_LIBELLE').AsString);
       End;
end;

procedure TOF_PGEM_MULCATCURSUS.AjouterFavoris(Sender : TObject);
Var Q : TQuery;
    TobCursus : Tob;
    Cursus : String;
    i,Imax : Integer;
    Liste  : THGrid;
begin
    Liste := THGrid(GetControl('FLISTE'));
    if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
    begin
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
{$ENDIF}      
      Cursus := TFmul(Ecran).Q.FindField('PCU_CURSUS').asstring;
      If Not existeSQL('SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_CODESTAGE="'+Cursus+'" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'" AND PCC_MILLESIME="'+MillesimeEC+'"') then
      begin
          Q := OpenSQL('SELECT MAX(PCC_RANGCURSUS) FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_CODESTAGE="'+Cursus+'"',True);
          if Not Q.EOF then
          begin
              IMax := Q.Fields[0].AsInteger;
              if IMax <> 0 then IMax := IMax + 1
              else IMax := 1;
          end
          else IMax := 1 ;
          Ferme(Q) ;
          TobCursus := Tob.Create('CURSUSSTAGE',Nil,-1);
          TobCursus.PutValue('PCC_CURSUS','-C-');
          TobCursus.PutValue('PCC_LIBELLE',V_PGI.UserSalarie);
          TobCursus.PutValue('PCC_ORDRE',-1);
          TobCursus.PutValue('PCC_RANGCURSUS',Imax);
          TobCursus.PutValue('PCC_CODESTAGE',Cursus);
          TobCursus.PutValue('PCC_MILLESIME',MillesimeEC);
          TobCursus.InsertDB(Nil);
          TobCursus.Free;
      end;
    end;
  end
  else if liste.AllSelected then
  begin
     TFmul(Ecran).Q.First;
     while not TFmul(Ecran).Q.EOF do
     begin
{$IFDEF EAGLCLIENT}
         if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
{$ENDIF}
         Cursus := TFmul(Ecran).Q.FindField('PCU_CURSUS').AsString;
      If Not existeSQL('SELECT PCC_CURSUS FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_CODESTAGE="'+Cursus+'" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'" AND PCC_MILLESIME="'+MillesimeEC+'"') then
         begin
          Q := OpenSQL('SELECT MAX(PCC_ORDRE) FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_CODESTAGE="'+Cursus+'"',True);
          if Not Q.EOF then
          begin
              IMax := Q.Fields[0].AsInteger;
              if IMax <> 0 then IMax := IMax + 1
              else IMax := 1;
          end
          else IMax := 1 ;
          Ferme(Q) ;
          TobCursus := Tob.Create('CURSUSSTAGE',Nil,-1);
          TobCursus.PutValue('PCC_CURSUS','-C-');
          TobCursus.PutValue('PCC_LIBELLE',V_PGI.UserSalarie);
          TobCursus.PutValue('PCC_ORDRE',-1);
          TobCursus.PutValue('PCC_RANGCURSUS',Imax);
          TobCursus.PutValue('PCC_CODESTAGE',Cursus);
          TobCursus.PutValue('PCC_MILLESIME',MillesimeEC);
          TobCursus.InsertDB(Nil);
          TobCursus.Free;
         end;
         TFmul(Ecran).Q.Next;
    end;
  end;
  TFMul(Ecran).bSelectAll.Down := False;
  Liste.AllSelected := False;
  Liste.ClearSelected;
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{ TOF_PGEMCURSUS }

procedure TOF_PGEMCURSUS.OnArgument(S: String);
var Grille : THGrid; //PT4
    Bt     : TToolbarButton97;
begin
  inherited;
       ReadTokenSt(S);
       Cursus := ReadTokenSt(S);
       Rang   := ReadTokenSt(S);
       Ecran.Caption := Ecran.Caption + ' ' + ReadTokenSt(S);
       UpdateCaption(Ecran);

       //PT4 - Début
       Grille := THGrid(GetControl('GRSTAGES'));
       If Grille <> Nil Then Grille.OnDblClick := GrilleDblClick;
       //PT4 - Fin

       //PT5 - Début
       Bt := TToolBarButton97(GetControl('BFAVORIS'));
       If Bt <> Nil then bt .OnClick := AjouterAuxFavoris;
       Bt := TToolBarButton97(GetControl('BNONFAVORIS'));
       If Bt <> Nil then bt .OnClick := RetirerDesFavoris;
       //PT5 - Fin

       MillesimeEC := RendMillesimeEManager;
end;

procedure TOF_PGEMCURSUS.OnLoad;
var 
    T : TOB;
    Somme : Double;
    Grille : THGrid;
begin
  inherited;

  T := TOB.Create('LesStages', Nil, -1);
  T.LoadDetailFromSQL('SELECT PST_CODESTAGE, PST_LIBELLE, PST_DUREESTAGE, PST_COUTBUDGETE, PST_COUTUNITAIRE FROM STAGE '+
                      ' RIGHT JOIN CURSUSSTAGE ON PCC_CODESTAGE=PST_CODESTAGE AND PCC_MILLESIME=PST_MILLESIME '+
                      ' WHERE PCC_CURSUS="'+Cursus+'" AND PCC_MILLESIME="0000"');
  If T <> Nil Then
  Begin
       Grille := THGrid(GetControl('GRSTAGES')); //PT4
       Grille.ColWidths[0] := -1;        // On cache le code stage
       T.PutGridDetail(Grille, False, False, 'PST_CODESTAGE;PST_LIBELLE;PST_DUREESTAGE'); //PT3
       Somme := T.SommeSimple('PST_DUREESTAGE',[''],['']);
       SetControlText('NBHEURES', FloatToStr(Somme));
       Somme := T.SommeSimple('PST_COUTBUDGETE',[''],['']);
       SetControlText('COUTBUDGET', FloatToStr(Somme));
       Somme := T.SommeSimple('PST_COUTUNITAIRE',[''],['']);
       SetControlText('COUTFORFAIT', FloatToStr(Somme));
  End;
  FreeAndNil(T);

  //PT5 - Début
  If existeSQL('SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.userSalarie+'" AND PCC_CODESTAGE="'+Cursus+'"') then
  Begin
       SetControlEnabled('BFAVORIS',False);
       SetControlEnabled('BNONFAVORIS',True);
  End
  Else
  Begin
       SetControlEnabled('BFAVORIS',True);
       SetControlEnabled('BNONFAVORIS',False);
  End;
  //PT5 - Fin
end;

//PT4 - Début
Procedure TOF_PGEMCURSUS.GrilleDblClick (Sender : TObject);
var Grille : THGrid;
    Stage  : String;
begin
       Grille := THGrid(GetControl('GRSTAGES'));
       If Grille <> Nil Then
       Begin
            Stage := Grille.CellValues[0, Grille.Row];
            If Stage <> '' Then
               AglLanceFiche('PAY','EM_CATALOGUE','',Stage+';0000','CONSULTCAT')
       End;
End;
//PT4 - Fin

//PT5 - Début
procedure TOF_PGEMCURSUS.AjouterAuxFavoris(Sender : Tobject);
var Q : TQuery;
    TobCursus : Tob;
    IMax : Integer;
begin
     Q := OpenSQL('SELECT MAX(PCC_ORDRE) FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_CODESTAGE="'+Cursus+'"',True);
     if Not Q.EOF then
     begin
          IMax := Q.Fields[0].AsInteger;
          if IMax <> 0 then IMax := IMax + 1
          else IMax := 1;
     end
     else IMax := 1 ;
     Ferme(Q) ;
     TobCursus := Tob.Create('CURSUSSTAGE',Nil,-1);
     TobCursus.PutValue('PCC_CURSUS','-C-');
     TobCursus.PutValue('PCC_LIBELLE',V_PGI.UserSalarie);
     TobCursus.PutValue('PCC_ORDRE',-1);
     TobCursus.PutValue('PCC_RANGCURSUS',Imax);
     TobCursus.PutValue('PCC_CODESTAGE',Cursus);
     TobCursus.PutValue('PCC_MILLESIME',MillesimeEC);
     TobCursus.InsertDB(Nil);
     SetControlEnabled('BFAVORIS',False);
     SetControlEnabled('BNONFAVORIS',True);
     TobCursus.Free;
end;

procedure TOF_PGEMCURSUS.RetirerDesFavoris(Sender : Tobject);
begin
    
     ExecuteSQL('DELETE FROM CURSUSSTAGE WHERE PCC_LIBELLE="'+V_PGI.UserSalarie+'" AND '+
     'PCC_CURSUS="-C-" AND PCC_MILLESIME="'+MillesimeEC+'" AND PCC_CODESTAGE="'+Cursus+'"');
     SetControlEnabled('BFAVORIS',True);
     SetControlEnabled('BNONFAVORIS',False);
end;
//PT5 - Fin
{$ENDIF}
//PT3 - Fin
Initialization
  registerclasses ( [ TOF_PGMUL_CURSUS,TOF_PGMUL_CURSUSEFOR,TOF_PGEDITCURSUS{$IFDEF EMANAGER},TOF_PGEM_MULCATCURSUS,TOF_PGEMCURSUS{$ENDIF} ] ) ;
end.

