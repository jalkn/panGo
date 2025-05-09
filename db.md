erDiagram
    Personas ||--|{ BienesyRentas : "posee"
    Personas ||--|{ DeclaracionConflictos : "declara"
    Personas ||--|{ LineaTransparencia : "ingresa"
    Personas ||--|{ AccesosSAP : "tiene"
    Personas ||--|{ SuccessFactor : "trabaja"
    Personas ||--o{ Alertas : "responsable"

    Bienes }|..|{ Alertas : "puede_generar"
    InformesAuditoria }|..|{ Alertas : "puede_generar"
    DeclaracionConflictos }|..|{ Alertas : "puede_generar"
    LineaTransparencia }|..|{ Alertas : "puede_generar"
    AccesosSAP }|..|{ Alertas : "puede_generar"
    SuccessFactor }|..|{ Alertas : "puede_generar"
    InformacionEspecial }|..|{ Alertas : "puede_generar"

    Personas {
        INT IDpersona PK "Primary Key"
        VARCHAR nombreCompleto
        VARCHAR cargo
        VARCHAR cedula "Unique Key"
        VARCHAR email
        VARCHAR compania
        VARCHAR Estado CHECK (Estado IN ('Activo', 'Retirado'))
        INT fkIdPersona FK 
        INT fkIdPeriodo FK 
    }

    BienesyRentas {
        INT idBien PK "Primary Key"
        VARCHAR Cedula
        VARCHAR Usuario
        VARCHAR Nombre
        VARCHAR Compania
        VARCHAR Cargo
        INT fkIdPeriodo FK "Foreign Key"
        INT Año_Declaracion
        INT Año_Creacion
        DECIMAL Activos
        INT Cant_Bienes
        INT Cant_Bancos
        INT Cant_Cuentas
        INT Cant_Inversiones
        DECIMAL Pasivos
        INT Cant_Deudas
        DECIMAL Patrimonio
        DECIMAL Apalancamiento
        DECIMAL Endeudamiento
        DECIMAL Activos_Var_Abs
        DECIMAL Activos_Var_Rel
        DECIMAL Pasivos_Var_Abs
        DECIMAL Pasivos_Var_Rel
        DECIMAL Patrimonio_Var_Abs
        DECIMAL Patrimonio_Var_Rel
        DECIMAL Apalancamiento_Var_Abs
        DECIMAL Apalancamiento_Var_Rel
        DECIMAL Endeudamiento_Var_Abs
        DECIMAL Endeudamiento_Var_Rel
        DECIMAL BancoSaldo
        DECIMAL Bienes
        DECIMAL Inversiones
        DECIMAL BancoSaldo_Var_Abs
        DECIMAL BancoSaldo_Var_Rel
        DECIMAL Bienes_Var_Abs
        DECIMAL Bienes_Var_Rel
        DECIMAL Inversiones_Var_Abs
        DECIMAL Inversiones_Var_Rel
        DECIMAL Ingresos
        INT Cant_Ingresos
        DECIMAL Ingresos_Var_Abs
        DECIMAL Ingresos_Var_Rel
    }

    Periodos {
        INT idPeriodo PK "Primary Key"
        INT Año_Declaración
        INT Año_Creación
        DATE Fecha_Inicio 
        DATE Fecha_Salida
    }

    Conflictos {
        INT ID PK "Primary Key"
        VARCHAR Cedula
        VARCHAR Nombre
        VARCHAR Compania
        VARCHAR Cargo
        VARCHAR Email
        DATE Fecha_de_Inicio
        BOOLEAN Q1
        BOOLEAN Q2
        BOOLEAN Q3
        BOOLEAN Q4
        BOOLEAN Q5
        BOOLEAN Q6
        BOOLEAN Q7
        BOOLEAN Q8
        BOOLEAN Q9
        BOOLEAN Q10
    }

    TarjetasMasterCard {
        VARCHAR Archivo
        VARCHAR Tarjetahabiente
        VARCHAR Número de Tarjeta
        VARCHAR Moneda
        DECIMAL Tipo de Cambio
        VARCHAR Número de Autorización
        DATE Fecha de Transacción
        TEXT Descripción
        DECIMAL Valor Original
        DECIMAL Tasa Pactada
        DECIMAL Tasa EA Facturada
        DECIMAL Cargos y Abonos
        DECIMAL Saldo a Diferir
        VARCHAR Cuotas
        INT Página
    }

    TransaccionesVisa {
        VARCHAR Archivo
        VARCHAR Tarjetahabiente
        VARCHAR Número de Tarjeta
        VARCHAR Número de Autorización
        DATE Fecha de Transacción
        TEXT Descripción
        DECIMAL Valor Original
        DECIMAL Tasa Pactada
        DECIMAL Tasa EA Facturada
        DECIMAL Cargos y Abonos
        DECIMAL Saldo a Diferir
        VARCHAR Cuotas
        INT Página
    }

    InformesAuditoria {
        INT IDinforme PK "Primary Key"
        DATE fechaInforme
        TEXT descripcion
        VARCHAR tipoInforme
        VARCHAR estado
    }

    LineaTransparencia {
        INT idEntrada PK "Primary Key"
        INT personaEntraID FK "Foreign Key"
        DATE fechaEntrada
        TEXT descripcion
        VARCHAR tipoInformacion
        VARCHAR estado
    }

    AccesosSAP {
        INT idSAP PK "Primary Key"
        INT personaSAPid FK "Foreign Key"
        VARCHAR usuario
        VARCHAR rol
        DATE fecha_creacion
        DATE fecha_modificacion
        VARCHAR estado
    }

    SuccessFactor {
        INT idFactor PK "Primary Key"
        INT personaFactorID FK "Foreign Key"
        VARCHAR puesto
        VARCHAR departamento
        DATE fecha_ingreso
    }

    InformacionEspecial {
        INT idInforme PK "Primary Key"
        DATE fecha
        TEXT descripcion
        VARCHAR tipo
    }

    Alertas {
        INT idAlerta PK "Primary Key"
        DATETIME fechaDeteccion
        VARCHAR tipo
        TEXT descripcion
        VARCHAR Estado CHECK (Estado IN ('Bajo', 'Medio', 'Alto'))
        VARCHAR tablaOrigen
        INT idTablaOrigen
        VARCHAR Estado CHECK (Estado IN ('Pendiente', 'En progreso', 'Resuelto'))
        VARCHAR usuarioResponsable
        INT personaID FK "Foreign Key"
    }