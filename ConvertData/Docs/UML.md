# UML Диаграммы ConvertData

## 📐 Диаграмма классов (Class Diagram)

### PlantUML код

```plantuml
@startuml ConvertData_ClassDiagram

' Стили
skinparam classAttributeIconSize 0
skinparam class {
    BackgroundColor<<Application>> LightBlue
    BackgroundColor<<Domain>> LightYellow
    BackgroundColor<<Infrastructure>> LightGreen
    BackgroundColor<<Entity>> LightCyan
}

' ========== Presentation Layer ==========
package "Presentation" {
    class Program {
        {static} +Main(args: string[]): void
    }
}

' ========== Application Layer ==========
package "Application" <<Application>> {
    class ConvertApp {
        -_writer: IRowWriter
        -_readerFactory: IRowReaderFactory
        -_pathResolver: IPathResolver
        -_licenseConfig: ILicenseConfigurator
        +Run(args: string[]): void
        -CreateJsonFromExcel(...): void
        -ApplyProfilesFromExcel(...): void
        -MergeAllAndExport(...): void
    }
    
    class JsonRecordEnricher {
        {static} -GeometrySubSections: string[]
        +Enrich(arr: JsonArray): int
        {static} -EnrichRecord(...): void
        {static} -DeepCopyNode(...): void
        {static} -ScoreCompleteness(...): int
        {static} -CountNonZeroValues(...): int
        {static} -GetNumericValue(...): double
    }
    
    class ConnectionCodeDeduplicator {
        +CreateDeduplicatedJson(...): int
        {static} -ExtractPrefix(...): string?
    }
    
    class ProfileExcelToJsonExporter {
        {static} -FileCategoryMap: Dictionary
        +Export(...): void
    }
    
    class ProfileLookupLoader {
        +Load(dir: string): Dictionary<string, ProfileGeometry>
    }
    
    class JsonProfilePatcher {
        +ApplyProfilesToJson(...): void
    }
    
    class RunModeParser {
        {static} +GetMode(args: string[]): RunMode
        {static} +GetInputArgsForCreateJson(...): string[]
        {static} +GetProfileColumn(...): string?
    }
    
    enum RunMode {
        All
        CreateJson
        ApplyProfiles
    }
    
    interface IRowReader {
        +Read(path: string): List<Row>
    }
    
    interface IRowWriter {
        +Write(rows: List<Row>, outputPath: string): void
    }
    
    interface IRowReaderFactory {
        +Create(path: string): IRowReader
    }
    
    interface IPathResolver {
        +GetProjectDir(startDir: string): string?
    }
    
    interface ILicenseConfigurator {
        +Configure(): void
    }
}

' ========== Domain Layer ==========
package "Domain" <<Domain>> {
    class Row {
        +Name: string
        +CONNECTION_CODE: string
        +variable: int
        +TableBrand: string
        
        +ProfileBeam: string
        +Beam_H: double
        +Beam_B: double
        +Beam_s: double
        +Beam_t: double
        +Beam_A: double
        +Beam_P: double
        +Beam_Iz: double
        +Beam_Iy: double
        +Beam_Ix: double
        +Beam_Wz: double
        +Beam_Wy: double
        +Beam_Wx: double
        +Beam_Sz: double
        +Beam_Sy: double
        +Beam_iz: double
        +Beam_iy: double
        +Beam_xo: double
        +Beam_yo: double
        
        +ProfileColumn: string
        +Column_H: double
        ' ... (аналогично для Column)
        
        +Plate_H: double
        +Plate_B: double
        +Plate_t: double
        
        +Flange_Lb: double
        +Flange_H: double
        +Flange_B: double
        +Flange_t: double
        
        +Stiff_tbp: double
        ' ... (остальные Stiff)
        
        +CoordinatesBolts: List<CoordinatesBolts>
        +F: int
        +Bolts_Nb: int
        +N_Rows: int
        +OptionBolts: double
        +e1: int
        +d1: int
        +d2: int
        +p1..p10: double
        
        +kf1..kf10: int
        
        +Sj: int
        +Sjo: int
        
        +Nt: int
        +Nc: int
        +N: int
        +Qy: int
        +Qz: int
        +Qx: int
        +My: int
        +T: int
        +Mneg: double
        +Mz: double
        +Mx: double
        +Mw: double
        
        +Alpha: double
        +Beta: double
        +Gamma: double
        +Delta: double
        +Epsilon: double
        +Lambda: double
    }
    
    class CoordinatesBolts {
        +X: int
        +Y: int
        +Z: int
        +CoordinatesBolts(x: int, y: int, z: int)
    }
    
    class ProfileGeometry <<record>> {
        +H: double
        +B: double
        +t_w: double
        +t_f: double
        +r1: double
        +r2: double
        +A: double
        +P: double
        +Iz: double
        +Iy: double
        +Ix: double
        +Wz: double
        +Wy: double
        +Wx: double
        +Sz: double
        +Sy: double
        +iz: double
        +iy: double
        +iu: double
        +xo: double
        +yo: double
    }
}

' ========== Infrastructure Layer ==========
package "Infrastructure" <<Infrastructure>> {
    class EpplusRowReader {
        +Read(path: string): List<Row>
        {static} -ReadXlsxWithEpplus(...): List<Row>
        {static} -GetCell(...): string
        {static} -FindHeaderRow(...): int
        {static} -MergeAdditionalSheets(...): void
        {static} -MergeSheet(...): void
        {static} -EnsureBolts(...): void
        {static} -BuildBoltsColumnMap(): Dictionary
    }
    
    class JsonRowWriter {
        {static} -BoltYKeys: string[]
        {static} -BoltXKeys: string[]
        +Write(...): void
        {static} -WriteBeam(...): void
        {static} -WriteColumn(...): void
        {static} -WritePlate(...): void
        {static} -WriteFlange(...): void
        {static} -WriteStiff(...): void
        {static} -WriteBolts(...): void
        {static} -WriteBoltY(...): void
        {static} -WriteBoltX(...): void
        {static} -WriteBoltZ(...): void
        {static} -WriteWelds(...): void
        {static} -WriteInternalForces(...): void
        {static} -Dbl(v: double): string
        {static} -JsonEscape(s: string): string
    }
    
    class RowReaderFactory {
        +Create(path: string): IRowReader
    }
    
    class PathResolver {
        +GetProjectDir(startDir: string): string?
        {static} +HasExcelExtension(path: string): bool
    }
    
    class EpplusLicenseConfigurator {
        +Configure(): void
    }
    
    class ExcelColumnMap {
        +IdxH: int
        +IdxB: int
        +Idxs: int
        +Idxt: int
        +IdxName: int
        +IdxCode: int
        +IdxProfile: int
        ' ... (все остальные индексы)
        +IdxAlpha..IdxLambda: int
        +IsMainTable: bool <<get>>
        +IsProfileTable: bool <<get>>
    }
    
    class ExcelHeaderResolver {
        {static} +ProfileColumnOverride: string?
        {static} +Resolve(header: List<string>): ExcelColumnMap
        {static} -ResolveGreekFallback(...): void
        {static} +ApplyProfileFallback(...): void
    }
    
    class RowMapper {
        {static} +MapMainRow(...): Row
        {static} +MapProfileRow(...): Row
    }
    
    class JsonMerger {
        +MergeAll(jsonDir: string): JsonArray
    }
    
    package "Parsing" {
        class NumericParser {
            {static} -RuCulture: CultureInfo
            {static} +ParseDouble(s: string?): double
            {static} +ParseInt(s: string?): int
        }
        
        class HeaderUtils {
            {static} +NormalizeHeader(h: string): string
            {static} +IndexOfHeader(...): int
            {static} +IndexOfHeaderAny(...): int
        }
        
        enum ExcelFileFormat {
            Unknown
            ZipXlsx
            CompoundFileBinary
        }
        
        class ExcelFileSignature {
            {static} +Detect(path: string): ExcelFileFormat
        }
    }
    
    package "Interop" {
        class ExcelXlsConverter {
            {static} +ConvertXlsToXlsxViaExcel(...): void
        }
    }
}

' ========== Entities ==========
package "Entitys" <<Entity>> {
    class ConnectionCodeItem {
        +CONNECTION_GUID: Guid
        +CONNECTION_CODE: string
    }
    
    class ProfileItem {
        +CONNECTION_GUID: Guid
        +Profile: string
    }
    
    class NameItem {
        +NAME_GUID: Guid
        +Name: string
    }
}

' ========== Relationships ==========

' Program → ConvertApp
Program ..> ConvertApp : creates

' ConvertApp dependencies
ConvertApp --> IRowWriter : uses
ConvertApp --> IRowReaderFactory : uses
ConvertApp --> IPathResolver : uses
ConvertApp --> ILicenseConfigurator : uses
ConvertApp ..> JsonRecordEnricher : uses
ConvertApp ..> ConnectionCodeDeduplicator : uses
ConvertApp ..> ProfileExcelToJsonExporter : uses
ConvertApp ..> ProfileLookupLoader : uses
ConvertApp ..> JsonProfilePatcher : uses
ConvertApp ..> RunModeParser : uses
ConvertApp ..> JsonMerger : uses

' Interface implementations
EpplusRowReader ..|> IRowReader
JsonRowWriter ..|> IRowWriter
RowReaderFactory ..|> IRowReaderFactory
PathResolver ..|> IPathResolver
EpplusLicenseConfigurator ..|> ILicenseConfigurator

' RowReaderFactory → EpplusRowReader
RowReaderFactory ..> EpplusRowReader : creates

' EpplusRowReader dependencies
EpplusRowReader --> ExcelHeaderResolver : uses
EpplusRowReader --> RowMapper : uses
EpplusRowReader --> NumericParser : uses
EpplusRowReader --> HeaderUtils : uses
EpplusRowReader --> ExcelFileSignature : uses
EpplusRowReader --> ExcelXlsConverter : uses
EpplusRowReader --> CoordinatesBolts : creates
EpplusRowReader --> Row : creates

' RowMapper → NumericParser
RowMapper --> NumericParser : uses
RowMapper --> Row : creates

' JsonRowWriter → Row
JsonRowWriter --> Row : reads

' ProfileLookupLoader → ProfileGeometry
ProfileLookupLoader --> ProfileGeometry : creates

' Row → CoordinatesBolts
Row "1" o-- "*" CoordinatesBolts : contains

@enduml
```

### Визуализация

Для рендеринга этой диаграммы используйте:
- [PlantUML Online Editor](https://www.plantuml.com/plantuml/uml/)
- VS Code расширение "PlantUML"
- IntelliJ IDEA PlantUML integration

---

## 📊 Диаграмма последовательности (Sequence Diagram)

### Сценарий: Полный цикл конвертации

```plantuml
@startuml ConvertData_SequenceDiagram

actor User
participant Program
participant ConvertApp
participant RunModeParser
participant RowReaderFactory
participant EpplusRowReader
participant ExcelXlsConverter
participant ExcelHeaderResolver
participant RowMapper
participant NumericParser
participant JsonRowWriter
participant ProfileLookupLoader
participant JsonProfilePatcher
participant JsonMerger
participant JsonRecordEnricher
participant ConnectionCodeDeduplicator
database Excel
database JSON

title Полный цикл конвертации ConvertData

User -> Program : запуск ConvertData.exe
activate Program

Program -> ConvertApp : new ConvertApp()
activate ConvertApp

Program -> ConvertApp : Run(args)

ConvertApp -> RunModeParser : GetMode(args)
activate RunModeParser
RunModeParser --> ConvertApp : RunMode.All
deactivate RunModeParser

== Этап 1: CreateJson ==

loop для каждого .xls/.xlsx файла в EXCEL/
    ConvertApp -> RowReaderFactory : Create(path)
    activate RowReaderFactory
    
    RowReaderFactory -> EpplusRowReader : new EpplusRowReader()
    activate EpplusRowReader
    RowReaderFactory --> ConvertApp : IRowReader
    deactivate RowReaderFactory
    
    ConvertApp -> EpplusRowReader : Read(path)
    
    EpplusRowReader -> Excel : определить формат
    activate Excel
    Excel --> EpplusRowReader : .xls (CompoundFileBinary)
    deactivate Excel
    
    EpplusRowReader -> ExcelXlsConverter : ConvertXlsToXlsxViaExcel(xls, tmpXlsx)
    activate ExcelXlsConverter
    ExcelXlsConverter -> Excel : COM Interop: Open + SaveAs
    activate Excel
    Excel --> ExcelXlsConverter : tmpXlsx создан
    deactivate Excel
    ExcelXlsConverter --> EpplusRowReader
    deactivate ExcelXlsConverter
    
    EpplusRowReader -> Excel : ReadXlsxWithEpplus(tmpXlsx)
    activate Excel
    
    EpplusRowReader -> ExcelHeaderResolver : Resolve(header)
    activate ExcelHeaderResolver
    ExcelHeaderResolver --> EpplusRowReader : ExcelColumnMap
    deactivate ExcelHeaderResolver
    
    loop для каждой строки данных
        EpplusRowReader -> RowMapper : MapMainRow(...)
        activate RowMapper
        
        loop для каждого поля
            RowMapper -> NumericParser : ParseDouble/ParseInt(...)
            activate NumericParser
            NumericParser --> RowMapper : значение
            deactivate NumericParser
        end
        
        RowMapper --> EpplusRowReader : Row
        deactivate RowMapper
    end
    
    EpplusRowReader -> EpplusRowReader : MergeAdditionalSheets\n(geometry, bolts, weld)
    
    Excel --> EpplusRowReader : List<Row>
    deactivate Excel
    
    EpplusRowReader --> ConvertApp : List<Row>
    deactivate EpplusRowReader
    
    ConvertApp -> JsonRowWriter : Write(rows, JSON_OUT/file.json)
    activate JsonRowWriter
    JsonRowWriter -> JSON : сохранить
    activate JSON
    JSON --> JsonRowWriter
    deactivate JSON
    JsonRowWriter --> ConvertApp
    deactivate JsonRowWriter
end

== Этап 2: ApplyProfiles ==

ConvertApp -> ProfileLookupLoader : Load(EXCEL_Profile/)
activate ProfileLookupLoader
ProfileLookupLoader -> Excel : читать ProfileBeam.xls
activate Excel
Excel --> ProfileLookupLoader : Dictionary<string, ProfileGeometry>
deactivate Excel
ProfileLookupLoader --> ConvertApp : lookup
deactivate ProfileLookupLoader

ConvertApp -> JsonProfilePatcher : ApplyProfilesToJson(JSON_OUT/, lookup)
activate JsonProfilePatcher

loop для каждого .json в JSON_OUT/
    JsonProfilePatcher -> JSON : прочитать
    activate JSON
    JSON --> JsonProfilePatcher : JsonArray
    deactivate JSON
    
    loop для каждого объекта
        JsonProfilePatcher -> JsonProfilePatcher : найти профиль в lookup
        JsonProfilePatcher -> JsonProfilePatcher : обновить Beam_H, Beam_B, ...
    end
    
    JsonProfilePatcher -> JSON : сохранить
    activate JSON
    JSON --> JsonProfilePatcher
    deactivate JSON
end

JsonProfilePatcher --> ConvertApp
deactivate JsonProfilePatcher

== Этап 3: Merge ==

ConvertApp -> JsonMerger : MergeAll(JSON_OUT/)
activate JsonMerger

loop для каждого .json в JSON_OUT/
    JsonMerger -> JSON : прочитать
    activate JSON
    JSON --> JsonMerger : JsonArray
    deactivate JSON
end

JsonMerger --> ConvertApp : merged JsonArray
deactivate JsonMerger

== Этап 3.5: Enrich ==

ConvertApp -> JsonRecordEnricher : Enrich(merged)
activate JsonRecordEnricher

JsonRecordEnricher -> JsonRecordEnricher : группировать по CONNECTION_CODE
JsonRecordEnricher -> JsonRecordEnricher : найти наиболее полные записи

loop для каждой группы
    loop для каждой неполной записи
        JsonRecordEnricher -> JsonRecordEnricher : DeepCopyNode(Geometry)
        JsonRecordEnricher -> JsonRecordEnricher : DeepCopyNode(Bolts)
        JsonRecordEnricher -> JsonRecordEnricher : DeepCopyNode(Welds)
        JsonRecordEnricher -> JsonRecordEnricher : копировать TableBrand
    end
end

JsonRecordEnricher --> ConvertApp : количество обогащённых
deactivate JsonRecordEnricher

ConvertApp -> JSON : сохранить all.json
activate JSON
JSON --> ConvertApp
deactivate JSON

== Этап 7: Deduplicate ==

ConvertApp -> ConnectionCodeDeduplicator : CreateDeduplicatedJson(...)
activate ConnectionCodeDeduplicator

ConnectionCodeDeduplicator -> JSON : прочитать all.json
activate JSON
JSON --> ConnectionCodeDeduplicator : JsonArray
deactivate JSON

ConnectionCodeDeduplicator -> ConnectionCodeDeduplicator : найти дубликаты
ConnectionCodeDeduplicator -> ConnectionCodeDeduplicator : переименовать: M-3 → M-6, M-7

ConnectionCodeDeduplicator -> JSON : сохранить all_NotDuplicate.json
activate JSON
JSON --> ConnectionCodeDeduplicator
deactivate JSON

ConnectionCodeDeduplicator --> ConvertApp : количество заменённых
deactivate ConnectionCodeDeduplicator

ConvertApp --> Program : завершено
deactivate ConvertApp

Program --> User : нажмите клавишу...
deactivate Program

@enduml
```

---

## 🏛️ Компонентная диаграмма (Component Diagram)

```plantuml
@startuml ConvertData_ComponentDiagram

skinparam component {
    BackgroundColor<<Application>> LightBlue
    BackgroundColor<<Domain>> LightYellow
    BackgroundColor<<Infrastructure>> LightGreen
}

package "ConvertData Application" {
    component "Presentation Layer" {
        [Program] as Program
    }
    
    component "Application Layer" <<Application>> {
        [ConvertApp] as ConvertApp
        [JsonRecordEnricher] as Enricher
        [ConnectionCodeDeduplicator] as Dedup
        [ProfileExporter] as ProfileExporter
        [ProfilePatcher] as Patcher
        [ProfileLoader] as Loader
    }
    
    component "Domain Layer" <<Domain>> {
        [Row] as Row
        [CoordinatesBolts] as Bolts
        [ProfileGeometry] as ProfileGeom
        [Entities] as Entities
    }
    
    component "Infrastructure Layer" <<Infrastructure>> {
        [EpplusRowReader] as Reader
        [JsonRowWriter] as Writer
        [ExcelHeaderResolver] as HeaderResolver
        [RowMapper] as Mapper
        [JsonMerger] as Merger
        [NumericParser] as Parser
        [ExcelXlsConverter] as Converter
    }
    
    database "File System" {
        folder "EXCEL/" as ExcelIn
        folder "JSON_OUT/" as JsonOut
        folder "JSON_All/" as JsonAll
        folder "EXCEL_Profile/" as ProfileIn
        folder "EXCEL_Profile_OUT/" as ProfileOut
    }
    
    cloud "External Libraries" {
        [EPPlus 7.x] as EPPlus
        [System.Text.Json] as STJson
        [COM Interop Excel] as COMExcel
    }
}

' Connections
Program --> ConvertApp

ConvertApp --> Enricher
ConvertApp --> Dedup
ConvertApp --> ProfileExporter
ConvertApp --> Patcher
ConvertApp --> Loader
ConvertApp --> Reader
ConvertApp --> Writer
ConvertApp --> Merger

Reader --> Row : creates
Reader --> Bolts : creates
Reader --> HeaderResolver : uses
Reader --> Mapper : uses
Reader --> Parser : uses
Reader --> Converter : uses

Mapper --> Row : creates
Mapper --> Parser : uses

Writer --> Row : reads

Enricher --> STJson : uses
Dedup --> STJson : uses
Patcher --> STJson : uses
Merger --> STJson : uses

Loader --> ProfileGeom : creates

ProfileExporter --> Entities : creates

Reader --> EPPlus : uses
Converter --> COMExcel : uses

' File System
Reader ..> ExcelIn : reads
Writer ..> JsonOut : writes
Merger ..> JsonOut : reads
Merger ..> JsonAll : writes
Loader ..> ProfileIn : reads
ProfileExporter ..> ProfileOut : writes

@enduml
```

---

## 📦 Диаграмма пакетов (Package Diagram)

```plantuml
@startuml ConvertData_PackageDiagram

skinparam packageStyle rectangle

package ConvertData {
    package Presentation {
        class Program
    }
    
    package Application {
        package UseCases {
            class ConvertApp
            class JsonRecordEnricher
            class ConnectionCodeDeduplicator
            class ProfileExcelToJsonExporter
            class ProfileLookupLoader
            class JsonProfilePatcher
        }
        
        package Interfaces {
            interface IRowReader
            interface IRowWriter
            interface IRowReaderFactory
            interface IPathResolver
            interface ILicenseConfigurator
        }
        
        package Helpers {
            class RunModeParser
            enum RunMode
        }
    }
    
    package Domain {
        class Row
        class CoordinatesBolts
        class ProfileGeometry
    }
    
    package Infrastructure {
        package Readers {
            class EpplusRowReader
            class RowReaderFactory
        }
        
        package Writers {
            class JsonRowWriter
        }
        
        package Resolvers {
            class ExcelHeaderResolver
            class ExcelColumnMap
            class PathResolver
        }
        
        package Mappers {
            class RowMapper
            class JsonMerger
        }
        
        package Parsing {
            class NumericParser
            class HeaderUtils
            class ExcelFileSignature
        }
        
        package Interop {
            class ExcelXlsConverter
        }
        
        package Configuration {
            class EpplusLicenseConfigurator
        }
    }
    
    package Entitys {
        class ConnectionCodeItem
        class ProfileItem
        class NameItem
    }
}

' Dependencies
Presentation ..> Application
Application ..> Domain
Application ..> Infrastructure
Infrastructure ..> Domain

@enduml
```

---

## 🔄 Диаграмма активности (Activity Diagram)

### Процесс обогащения записей

```plantuml
@startuml JsonRecordEnricher_ActivityDiagram

|JsonRecordEnricher|
start

:Получить JsonArray;

if (Массив пустой?) then (да)
  :Вернуть 0;
  stop
endif

:Создать словарь групп\nпо CONNECTION_CODE;

|Группировка|
while (Для каждого объекта в массиве) is (есть ещё)
  :Извлечь CONNECTION_CODE;
  
  if (Код пустой?) then (да)
    :Пропустить;
  else (нет)
    :Добавить индекс в группу[код];
  endif
endwhile (все обработаны)

|Обогащение|
:enriched = 0;

while (Для каждой группы) is (есть ещё)
  if (В группе < 2 элементов?) then (да)
    :Пропустить группу;
  else (нет)
    
    |Поиск шаблона|
    :templateIdx = indices[0];
    :templateScore = ScoreCompleteness(arr[templateIdx]);
    
    while (Для k = 1..n) is (есть ещё)
      :score = ScoreCompleteness(arr[indices[k]]);
      
      if (score > templateScore?) then (да)
        :templateIdx = indices[k];
        :templateScore = score;
      endif
    endwhile (все проверены)
    
    if (templateScore == 0?) then (да)
      :Пропустить группу;
    else (нет)
      :template = arr[templateIdx];
      
      |Копирование|
      while (Для каждого idx в группе) is (есть ещё)
        if (idx == templateIdx?) then (да)
          :Пропустить (это шаблон);
        else (нет)
          :target = arr[idx];
          
          if (ScoreCompleteness(target) >= templateScore?) then (да)
            :Пропустить (уже полная);
          else (нет)
            :EnrichRecord(template, target);
            
            split
              :DeepCopy Geometry;
            split again
              :DeepCopy Bolts;
            split again
              :DeepCopy Welds;
            split again
              :Копировать TableBrand\n(если пустой);
            end split
            
            :enriched++;
          endif
        endif
      endwhile (все обработаны)
    endif
  endif
endwhile (все группы обработаны)

:Вернуть enriched;

stop

@enduml
```

---

## 📈 Диаграмма состояний (State Diagram)

### Жизненный цикл Row

```plantuml
@startuml Row_StateDiagram

[*] --> Created : new Row()

Created --> ReadFromExcel : EpplusRowReader.Read()

ReadFromExcel --> MainDataLoaded : основной лист прочитан

MainDataLoaded --> GeometryMerged : MergeSheet(geometry)
GeometryMerged --> BoltsMerged : MergeSheet(bolts)
BoltsMerged --> WeldsMerged : MergeSheet(weld)

WeldsMerged --> ProfileApplied : JsonProfilePatcher

ProfileApplied --> WrittenToJson : JsonRowWriter.Write()

WrittenToJson --> InMergedArray : JsonMerger.MergeAll()

InMergedArray --> Enriched : JsonRecordEnricher.Enrich()
note right
  Только для неполных записей
end note

Enriched --> Deduplicated : ConnectionCodeDeduplicator
note right
  CONNECTION_CODE может
  быть изменён
end note

Deduplicated --> [*] : сохранено в\nall_NotDuplicate.json

@enduml
```

---

## 🗺️ Диаграмма развертывания (Deployment Diagram)

```plantuml
@startuml ConvertData_DeploymentDiagram

node "Рабочая станция разработчика" {
    artifact "ConvertData.exe" as Exe
    
    folder "Проект ConvertData" {
        folder "EXCEL/" as ExcelFolder
        folder "JSON_OUT/" as JsonOutFolder
        folder "JSON_All/" as JsonAllFolder
        folder "EXCEL_Profile/" as ProfileFolder
        folder "EXCEL_Profile_OUT/" as ProfileOutFolder
    }
    
    component "Microsoft Excel" as Excel {
        note right
          Требуется для конвертации
          .xls → .xlsx через COM
        end note
    }
}

cloud "Зависимости NuGet" {
    artifact "EPPlus 7.x"
    artifact "System.Text.Json"
}

Exe --> ExcelFolder : читает *.xls, *.xlsx
Exe --> JsonOutFolder : пишет *.json
Exe --> JsonAllFolder : пишет all.json,\nall_NotDuplicate.json
Exe --> ProfileFolder : читает ProfileBeam.xls
Exe --> ProfileOutFolder : пишет Profile.json,\nCONNECTION_CODE.json

Exe ..> Excel : COM Interop\n(при наличии .xls)
Exe ..> "EPPlus 7.x" : использует
Exe ..> "System.Text.Json" : использует

@enduml
```

---

## 🔍 Диаграмма вариантов использования (Use Case Diagram)

```plantuml
@startuml ConvertData_UseCaseDiagram

left to right direction

actor "Инженер-конструктор" as Engineer
actor "Система расчёта" as CalcSystem

rectangle "ConvertData System" {
    usecase "Конвертировать\nExcel → JSON" as UC1
    usecase "Применить справочник\nпрофилей" as UC2
    usecase "Объединить JSON файлы" as UC3
    usecase "Обогатить неполные\nзаписи" as UC4
    usecase "Устранить дубликаты\nCONNECTION_CODE" as UC5
    usecase "Экспортировать\nсправочники" as UC6
    usecase "Создать отчёты" as UC7
    
    UC1 .> UC2 : <<include>>
    UC3 .> UC4 : <<include>>
    UC3 .> UC5 : <<include>>
    UC3 .> UC6 : <<include>>
    UC3 .> UC7 : <<include>>
}

Engineer --> UC1
Engineer --> UC2
Engineer --> UC3

CalcSystem --> UC6 : использует\nсправочники

@enduml
```

---

## 📋 Диаграмма объектов (Object Diagram)

### Пример структуры данных после обогащения

```plantuml
@startuml ConvertData_ObjectDiagram

object "Row #1" as R1 {
    CONNECTION_CODE = "M-3"
    Name = "Балка-Колонна M"
    ProfileBeam = "20Б1"
    variable = 1
    TableBrand = "Т1"
    Beam_H = 200
    Beam_B = 100
    Plate_H = 300
    Plate_B = 150
    F = 20
    e1 = 50
    kf1 = 6
    Nt = 100
    My = 50
    Alpha = 1.0
}

object "Row #2" as R2 {
    CONNECTION_CODE = "M-3"
    Name = "Балка-Колонна M"
    ProfileBeam = "20Б1"
    variable = 2
    TableBrand = "Т1" ← скопировано
    Beam_H = 200
    Beam_B = 100
    Plate_H = 300 ← скопировано
    Plate_B = 150 ← скопировано
    F = 20 ← скопировано
    e1 = 50 ← скопировано
    kf1 = 6 ← скопировано
    Nt = 200 ← уникальное
    My = 80 ← уникальное
    Alpha = 1.1 ← уникальное
}

object "CoordinatesBolts #1" as CB1 {
    X = 60
    Y = 50
    Z = 0
}

object "CoordinatesBolts #2" as CB2 {
    X = 120
    Y = 50
    Z = 0
}

object "ProfileGeometry: 20Б1" as PG {
    H = 200
    B = 100
    t_w = 5.6
    t_f = 8.5
    A = 30.6
    Iz = 1840
    Iy = 155
}

R1 o-- CB1
R1 o-- CB2
R2 o-- CB1 : shares
R2 o-- CB2 : shares

R1 ..> PG : профиль применён
R2 ..> PG : профиль применён

note right of R2
  Обогащённая запись:
  - Geometry скопирована из R1
  - Bolts скопированы из R1
  - Welds скопированы из R1
  - TableBrand скопирован из R1
  - InternalForces уникальные
  - Coefficients уникальные
end note

@enduml
```

---

## 🛠️ Как использовать диаграммы

### Рендеринг PlantUML

1. **Online:**
   - [PlantUML Web Server](https://www.plantuml.com/plantuml/uml/)
   - Скопируйте код диаграммы и вставьте в редактор

2. **VS Code:**
   ```bash
   # Установить расширение
   ext install jebbs.plantuml
   
   # Установить PlantUML локально (требуется Java)
   # или использовать онлайн сервер
   ```

3. **IntelliJ IDEA:**
   - Встроенная поддержка PlantUML
   - Settings → Plugins → PlantUML Integration

4. **Командная строка:**
   ```bash
   java -jar plantuml.jar diagram.puml
   # Создаст diagram.png
   ```

### Экспорт изображений

PlantUML поддерживает несколько форматов:
- PNG (по умолчанию)
- SVG (векторный, масштабируемый)
- EPS (для публикаций)
- PDF (для документации)

```bash
java -jar plantuml.jar -tsvg UML.puml
```

---

**См. также:**
- [README](README.md)
- [Архитектура](Architecture.md)
- [Поток данных](DataFlow.md)
