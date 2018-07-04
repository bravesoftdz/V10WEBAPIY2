{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 20/03/2002
Modifié le ... : 26/04/2002
Description .. : Source TOM de la TABLE : NATCPTE (NATCPTE)
Mots clefs ... : TOM;NATCPTE
*****************************************************************}
unit NATCPTE_TOM;

interface

uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    Utob,
    eFiche,
    maineagl,
    eFichList,
{$ELSE}
    db,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    Fiche,
    FE_main,
    FichList,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HEnt1,
    LibChpLi,
    CPNATCPTECODELG_TOF,
    UTOM;

procedure CPLanceFiche_SaisieTableLibre(S:string);
Procedure FicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);

type
    TOM_NATCPTE = class(TOM)
        procedure OnNewRecord; override;
        procedure OnDeleteRecord; override;
        procedure OnUpdateRecord; override;
        procedure OnArgument(S:string); override;
        procedure OnClose; override;
    private
        ModeConsult:Boolean; // est-on en mode consultation ?
        CodeTableLibre:string; // Code de la table a saisir
        CodeEnreg:string; // Code de l'enregistrement a afficher
        LongueurCode:Integer; // Nombre de caractère utilisé pour le code
        NewLongueur:Boolean; // Indicateur de saisie de la taille du code
        ListeCode : TStringList;
        procedure RecupereLongueur;

    end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HMsgBox;

const
    MessageListe:array[1..5] of string =
    (
        'Vous devez renseigner un code.',
        'Vous devez renseigner un libellé.',
        'Le code que vous avez saisi existe déjà. Vous devez le modifier.',
        'L''enregistrement est inaccessible',
        'Voulez-vous mettre à jour les comptes référençant ce code de table libre ?'
        );

//==============================================================================
Procedure FicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Var
	Param : String;
begin
	Case Comment of
		taCreat,taCreatOne : Param := 'ACTION=CREATION';
		taModif            : Param := 'ACTION=MODIFICATION';
	else Param := 'ACTION=CONSULTATION';
	end;

	Param := Param + ';' + Code + ';' + Compte;

	CPLanceFiche_SaisieTableLibre(Param);
END;

//==============================================================================
procedure CPLanceFiche_SaisieTableLibre(S:string);
var
	Args, Arg1, Arg2 : String;
begin
    // Récupération des paramètres de lancement
    Args := Copy(S, 1, Length(S));
    ReadTokenSt(S); // Action = ...
    Arg1 := ReadTokenSt(S); // code table
    Arg2 := ReadTokenSt(S); // lequel
    AGLLanceFiche('CP','CPNATCPTE',Arg1,Arg2,Args);
end;

//==============================================================================
procedure TOM_NATCPTE.OnNewRecord;
begin
    inherited;

    SetField('NT_TYPECPTE', CodeTableLibre);
    SetField('NT_DATE1', iDate1900 );
    SetField('NT_DATE2', iDate1900 );
    SetField('NT_DATE3', iDate1900 );
    SetField('NT_DATE4', iDate1900 );

    if Ecran <> nil then
    begin
        SetControlEnabled('NT_NATURE', True);
        SetFocusControl('NT_NATURE');
    end;
end;

procedure TOM_NATCPTE.OnDeleteRecord;
var
    Table, Champ, code:string;
begin
    inherited;

    // Effacement des champs référençant l'enregistrement à effacer
    code := GetField('NT_NATURE');
    if code <> '' then
    begin
        case CodeTableLibre[1] of
            'G':
                begin
                    Table := 'GENERAUX';
                    Champ := 'G_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'T':
                begin
                    Table := 'TIERS';
                    Champ := 'T_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'S':
                begin
                    Table := 'SECTION';
                    Champ := 'S_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'B':
                begin
                    Table := 'BUDGENE';
                    Champ := 'BG_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'D':
                begin
                    Table := 'BUDSECT';
                    Champ := 'BS_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'E':
                begin
                    Table := 'ECRITURE';
                    Champ := 'E_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'A':
                begin
                    Table := 'ANALYTIQ';
                    Champ := 'Y_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'U':
                begin
                    Table := 'BUDECR';
                    Champ := 'BE_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
            'I':
                begin
                    Table := 'IMMO';
                    Champ := 'I_TABLE' + Copy(CodeTableLibre, 3, 1);
                end;
        end;
        if ExisteSql('SELECT ' + Champ + ' FROM ' + Table + ' WHERE ' + Champ + '="' + Code + '"') then
            if PGIAsk(TraduireMemoire(MessageListe[5]), TraduireMemoire('Paramétrage des zones libres')) = mrYes then
                ExecuteSql('UPDATE ' + Table + ' SET ' + Champ + '="" ' + 'WHERE ' + Champ + '="' + Code + '"');
    end;
end;

procedure TOM_NATCPTE.OnUpdateRecord;
var
    code         : string;
    Ind          : integer;
begin
    inherited;
    // Code Obligatoire
    code := GetControlText('NT_NATURE');

    if code = '' then
    begin
        PGIBox(TraduireMemoire(MessageListe[1]), TraduireMemoire(Ecran.Caption));
        SetFocusControl('NT_NATURE');
        LastError := 1;
        Exit;
    end
    else if DS.State = dsInsert then
    begin
        if Length(code) > LongueurCode then
        code := copy(code,1,LongueurCode) ;
        if Length(code) < LongueurCode then
        code := BourreLaDoncSurLaTable(CodeTableLibre,code) ;
        { FQ 20807 BVE 12.07.07 }
        if Pos(#0,code) > 0 then
           SetLength(code,Pos(#0,code) - 1);
        { END FQ 20807 }
        SetControlText('NT_NATURE',code) ;
        // ajout me fiche 19189
        if ListeCode = nil then
        begin
            ListeCode := TStringList.Create;
            ListeCode.Duplicates := dupIgnore;
        end
        else
        begin
            Ind := ListeCode.Indexof (code);
            if Ind >= 0 then
            begin
                  PGIBox(TraduireMemoire(MessageListe[3]), TraduireMemoire(Ecran.Caption));
                  SetControlEnabled('NT_NATURE', True);
                  SetFocusControl('NT_NATURE');
                  LastError := 1;
                  Exit;
            end;
        end;
        If ExisteSQL('select * from NATCPTE where NT_TYPECPTE="'+CodeTableLibre+'" AND NT_NATURE="'+Code +'"') then
        begin
              PGIBox(TraduireMemoire(MessageListe[3]), TraduireMemoire(Ecran.Caption));
              SetControlEnabled('NT_NATURE', True);
              SetFocusControl('NT_NATURE');
              LastError := 1;
              Exit;
        end;
    end;

    // Libelle Obligatoire
    if GetControlText('NT_LIBELLE')='' then
    begin
        PGIBox(TraduireMemoire(MessageListe[2]), TraduireMemoire(Ecran.Caption));
        SetFocusControl('NT_LIBELLE');
        LastError := 1;
        Exit;
    end;

    // code non modifiable après validation...
    SetControlEnabled('NT_NATURE', False);
    if DS.State = dsInsert then ListeCode.Add(code);
end;

procedure TOM_NATCPTE.OnArgument(S:string);
begin
    inherited;

    NewLongueur := False;
    LongueurCode := 0;

    // Récupération des paramètres de lancement
    ModeConsult := pos('CONSULTATION', ReadTokenSt(S)) > 0; // Action = ...
    CodeTableLibre := ReadTokenSt(S);
    CodeEnreg := ReadTokenSt(S);

    // Initialisation interface
    if Ecran <> nil then
    begin
        TFFicheListe(Ecran).ModeleImpression := 'PRT_NATCPTE'; // FQ 12335
{$IFDEF CCS3}
        SetControlVisible('Pages', False);
{$ELSE}
        PersoChampsLibres(CodeTableLibre, TWinControl(GetControl('TInfos1')));
        PersoChampsLibres(CodeTableLibre, TWinControl(GetControl('TInfos2')));
{$ENDIF}
        // Si rien de visible dans les onglets d'infos, on les cache
        if (not GetControlVisible('TInfos1')) and (not GetControlVisible('TInfos2')) then SetControlVisible('Pages', False);
    end;

    // Récupérer ou saisie de la longueur max du code à saisir
    RecupereLongueur;
    SetControlProperty('NT_NATURE', 'MaxLength', LongueurCode);

end;

//==============================================================================
procedure TOM_NATCPTE.RecupereLongueur;
var
	i,j : Integer;
begin
    LongueurCode := 0;

    // Doit-on saisir la longueur du code ?
    NewLongueur := not Presence('NATCPTE', 'NT_TYPECPTE', CodeTableLibre);
    if NewLongueur and (not ModeConsult) then
    begin
    // Propose la saisie et mets à jour VH
        CPLanceFiche_CodeLongueur(CodeTableLibre);
    end;

    // Récupération de la longueur du code dans VH
    case CodeTableLibre[1] of
        'G' : i:=1 ;
        'T' : i:=2 ;
        'S' : i:=3 ;
        'B' : i:=4 ;
        'D' : i:=5 ;
        'E' : i:=6 ;
        'A' : i:=7 ;
        'U' : i:=8 ;
        'I' : i:=9 ;
    else i := 0;
    end;

    if i > 0 then
    begin
        j := StrToInt(Copy(CodeTableLibre, 2, 2)) + 1;
        LongueurCode := VH^.LgTableLibre[i, j];
    end;

    // si nouvelle longueur, MAJ du code si celui-ci est trop grand
    if NewLongueur then
    begin
        if (LongueurCode > 1) and (Length(CodeEnreg) > LongueurCode) then SetControlText('NT_NATURE', Copy(CodeEnreg, 1, LongueurCode));
        NewLongueur := False;
    end;
end;

procedure TOM_NATCPTE.OnClose ;
begin
  Inherited ;
 if ListeCode <> nil then
 begin
  ListeCode.free; ListeCode := nil;
 end;
end;
//==============================================================================
initialization
    registerclasses([TOM_NATCPTE]);
end.
