{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : REACREECOLIS ()
Mots clefs ... : TOF;REACREECOLIS; 
*****************************************************************}

unit REACREECOLIS_TOF ;

interface

uses StdCtrls,
     Controls,
     Classes,
     messages,
     windows,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     EdtEtat,
     EdtREtat,
{$ELSE}
     UtilEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Wtof,
     uTob,
     SaisieList,
     uTableFiltre;

type
  TOF_REACREECOLIS = Class (twTOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    Procedure TMP_NUMERO_Onexit(Sender : tObject);

    Private
      TF             : TTableFiltre;
      lClose         : Boolean;
      procedure DoSetNavigate(Sender:TObject);
      procedure bSelect_Onclick(Sender : tObject);
      procedure HNUMGROUP_OnExit(Sender : tObject);

    Public
      lPremier : Boolean;
  end ;

implementation

uses
  htb97,
  menus,
  Paramsoc,
  PropTransfFSL_TOF
  ;

const cNaturepiece = 'TRV';   // Nature de la piece de transfert de stock

procedure TOF_REACREECOLIS.HNUMGROUP_OnExit(Sender : tObject);
begin
  if not lCLose then
  begin
    lPremier := False;
    bSelect_Onclick(Sender);
  end;
end;


Procedure TOF_REACREECOLIS.bSelect_Onclick(Sender : tObject);

var
  cNumero : String;
begin
  if GetControl('HNUMGROUP') <> nil then
    cNumero := ThEdit(GetControl('HNUMGROUP')).Text
  Else  // Normalement, ca ne devrais jamais se passer, a moins que quelqu'un sabote la fiche...
    cNumero := '';

  // Pour appliquer une sélection sur la table 'fille', mise à jour de la propriété WhereTable
  // Puis on désactive puis réactive le tTableFiltre pour que ce where soit pris en compte.
  if (cNumero <> '') then
    TF.WhereTable := 'WHERE GP_NATUREPIECEG = "' + cNaturepiece + '" and GP_REFINTERNE = "' + cNumero + '"'
  else
  begin
    if lPremier then
      TF.WhereTable := 'WHERE GP_NATUREPIECEG = "' + cNaturepiece + '" and GP_REFINTERNE <> "' + cNumero + '" and GP_ETATEXPORT = "TRA"'
    else
      TF.WhereTable := 'WHERE GP_NATUREPIECEG = "/*-"'; //Pour ne rien voir
  end;

  TF.Active := False;
  TF.Active := True;
  ThSpinEdit(GetControl('TMP_NUMERO')).SetFocus;
end;


procedure TOF_REACREECOLIS.TMP_NUMERO_Onexit(Sender : tObject);
var
  iNumpiece : Integer;
  cNumero : String;
  q : tQuery;
  T : TOB;
	i: integer;
  lSortie: boolean;
begin
  if not lCLose then
  begin
    //cCart := GetString('TMP_ARTICLE');
    if GetControl('TMP_NUMERO') <> nil then
      iNumPiece := ThSpinEdit(GetControl('TMP_NUMERO')).Value
    else  // Normalement, ca ne devrais jamais se passer, a moins que quelqu'un sabote la fiche...
      iNumPiece := 0;

    if GetControl('HNUMGROUP') <> nil then
      cNumero := ThEdit(GetControl('HNUMGROUP')).Text
    else  // Normalement, ca ne devrais jamais se passer, a moins que quelqu'un sabote la fiche...
      cNumero := '';

    if (iNumPiece <> -1) and (iNumpiece <> 0) then
    begin
      // Recherche de l'enregistrement à mettre à jour dans la TOB de la fiche.
      T := TF.TobFiltre.FindFirst(['GP_REFINTERNE','GP_NATUREPIECEG','GP_NUMERO'],[cNumero,'TRV',iNumPiece],false);
      if T <> nil then
      begin
        i := TF.TobFiltre.Detail.indexOf( T )+1;
        TF.SelectRecord( i );

        // Mise à jours.
        TF.StartUpdate;
        TF.PutValue( 'GP_ETATEXPORT', 'ENV' );
        TF.PutValue( 'GP_DATEEXPORT', V_PGI.DateEntree );
        TF.Post;
        TF.EndUpdate;
      end
      else
      begin
        Ferme(Q);
        ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
        HShowMessage('0;ATTENTION : Cette pièce ne fait pas partie de ce groupage;;O;O;O;','','');
      end;
      ThSpinEdit(GetControl('TMP_NUMERO')).Value := 0;
      ThSpinEdit(GetControl('TMP_NUMERO')).SetFocus;
    end
    else
    begin
      if (iNumpiece = -1) and (cNumero <> '') then //Validation
      begin
        lSortie := not existeSql('SELECT GP_NUMERO '
                                +'FROM PIECE '
                                +'WHERE GP_REFINTERNE="'+cNumero+'" AND GP_NATUREPIECEG = "'+cNaturepiece+'"'
                                +'      AND GP_ETATEXPORT <> "ENV"');
        if not lSortie then
        begin
          ErreurSonore(GetParamSoc('SO_ERRFICSON'), TFSaisieList( ecran ));
          lSortie := (PgiAsk(TraduireMemoire('ATTENTION. Il reste des colis non scanné. Confirmez-vous la sortie du groupage ')+cNumero+' ?','') = mrYes);
        end;

        if lSortie then
        begin
{V500_004 Début}
          V_PGI.NoPrintDialog := True; //not BApercu à paramétrer pour avoir ou non la fenetre de dialogue impression
          if LanceEtat('E','LO2','LT4',False,False,False,Nil,'GP_REFINTERNE="'+cNumero+'"',TraduireMemoire('Fiche de colisage'),False) < 0 then
            V_PGI.IoError:=oeUnknown;
          V_PGI.NoPrintDialog := True; //not BApercu à paramétrer pour avoir ou non la fenetre de dialogue impression
          if LanceEtat('E','LO2','LT5',False,False,False,Nil,'GP_REFINTERNE="'+cNumero+'"',TraduireMemoire('Fiche de transfert'),False) < 0 then
            V_PGI.IoError:=oeUnknown;
{V500_004 Fin}
          ThSpinEdit(GetControl('TMP_NUMERO')).Value := 0;
          PostMessage(TWincontrol(GetControl('HNUMGROUP')).handle, WM_SETFOCUS, 0 , 0 );
          //ThEdit(GetControl('HNUMGROUP')).SetFocus;
        end
        else
        begin
          ThSpinEdit(GetControl('TMP_NUMERO')).Value := 0;
          ThSpinEdit(GetControl('TMP_NUMERO')).SetFocus;
        end;

      end;
    end;
  end;
end;



procedure TOF_REACREECOLIS.DoSetNavigate(Sender: TObject);
begin
  SetControlProperty('BTACTION', 'ENABLED', TF.CanReload);
  SetControlProperty('BTREE', 'ENABLED', TF.CanReload);
end;


procedure TOF_REACREECOLIS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REACREECOLIS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_REACREECOLIS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_REACREECOLIS.OnLoad ;
begin
  Inherited ;
  ThSpinEdit(GetControl('TMP_NUMERO')).Value := 0;
  thEdit(GetControl('HNUMGROUP')).SetFocus;
end ;

procedure TOF_REACREECOLIS.OnArgument (S : String ) ;
begin
  Inherited ;
  lClose := False;
  TF := TFSaisieList( ecran ).LeFiltre;
  TF.OnSetNavigate:= DoSetNavigate;
  // C'est un écran sur lequel on agis sur des lignes existantes. pas question de créer, modifier ou dupliquer
  tToolbarButton97(GetControl('bDelete')).Visible := False;
  tToolbarButton97(GetControl('bInsert')).Visible := False;
  tToolbarButton97(GetControl('bDupliquer')).Visible := False;
  //
  if GetControl('TMP_NUMERO')  <> nil then ThSpinEdit(GetControl('TMP_NUMERO')).OnExit := TMP_NUMERO_OnExit;
  if GetControl('bSelect')     <> nil then tButton(GetControl('bSelect')).Onclick := bSelect_Onclick;
  if GetControl('HNUMGROUP')   <> nil then thEdit(GetControl('HNUMGROUP')).OnExit := HNUMGROUP_OnExit;
  lPremier := True;
  bSelect_Onclick(Self);
  thEdit(GetControl('HNUMGROUP')).SetFocus;
end ;

procedure TOF_REACREECOLIS.OnClose ;
begin
  lCLose := True;
  Inherited ;
end ;

procedure TOF_REACREECOLIS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_REACREECOLIS.OnCancel () ;
begin
  Inherited ;
end ;

initialization
  registerclasses ( [ TOF_REACREECOLIS ] ) ;
end.
