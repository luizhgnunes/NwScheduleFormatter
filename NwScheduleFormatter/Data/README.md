# Data Layer Infrastructure

## Overview

The data layer was implemented using **Entity Framework Core 9.0** with **SQLite** as the database. The database is created automatically on the first application run.

## Directory Structure

```
NwScheduleFormatter/
├── Data/
│   ├── Entities/
│   │   ├── SendingOfDesignation.cs
│   │   └── DesignationContact.cs
│   ├── Enums/
│   │   └── ConfirmationSource.cs
│   ├── Repositories/
│   │   ├── IRepository.cs
│   │   ├── Repository.cs
│   │   ├── ISendingOfDesignationRepository.cs
│   │   ├── SendingOfDesignationRepository.cs
│   │   ├── IDesignationContactRepository.cs
│   │   └── DesignationContactRepository.cs
│   ├── NwScheduleFormatterDbContext.cs
│   └── RepositoryExtensions.cs
```

## Database Tables

### SendingOfDesignation Entity

Represents a WhatsApp designation sending, mapping the `tb_sending_of_designation` table:

| Property | SQL Column | Type | Notes |
|---|---|---|---|
| `Id` | `sod_id` | int | Primary Key, Auto-increment |
| `Reference` | `sod_reference` | varchar(10) | Unique |
| `DesignatedName` | `sod_designated_name` | varchar(100) | Nullable |
| `DesignationDate` | `sod_designation_date` | DATE | Date of designation |
| `CreationDate` | `sod_creation_date` | timestamp | Default: CURRENT_TIMESTAMP |
| `IsCanceled` | `sod_is_canceled` | boolean | Default: 0 |
| `WamId` | `sod_wamid` | varchar(200) | Nullable, WhatsApp message ID |
| `SendDate` | `sod_send_date` | timestamp | Nullable |
| `IsConfirmed` | `sod_is_confirmed` | boolean | Default: 0 |
| `ConfirmationDate` | `sod_confirmation_date` | timestamp | Nullable |
| `ConfirmationSource` | `sod_confirmation_source` | int8 | Default: 0, Confirmation source (enum) |
| `SendingError` | `sod_sending_error` | TEXT | Nullable, Error message if sending failed |

### DesignationContact Entity

Represents a designation contact, mapping the `tb_designation_contact` table:

| Property | SQL Column | Type | Notes |
|---|---|---|---|
| `Id` | `dc_id` | int | Primary Key, Auto-increment |
| `Name` | `dc_name` | varchar(100) | Contact name |
| `WhatsAppNumber` | `dc_whatsapp_number` | varchar(13) | WhatsApp phone number |
| `SkipSendingMessage` | `dc_skip_sending_message` | boolean | Default: 0, Skip sending to this contact |

## Enums

### ConfirmationSource

Represents the source of a designation sending confirmation:

```csharp
public enum ConfirmationSource : byte
{
	NotConfirmed = 0,  // Not yet confirmed
	WhatsApp = 1,      // Confirmed via WhatsApp
	Wpf = 2            // Confirmed via WPF application
}
```

## Repository Pattern

### Generic Interface `IRepository<T>`

```csharp
public interface IRepository<TEntity> where TEntity : class
{
	Task<IEnumerable<TEntity>> GetAllAsync();
	Task<TEntity?> GetByIdAsync(int id);
	Task<TEntity> AddAsync(TEntity entity);
	Task<TEntity> UpdateAsync(TEntity entity);
	Task DeleteAsync(int id);
	Task<bool> ExistsAsync(int id);
	Task SaveChangesAsync();
}
```

### SendingOfDesignation Repository

Interface: `ISendingOfDesignationRepository` extends `IRepository<SendingOfDesignation>`

Additional methods:
- `GetByReferenceAsync(string reference)` - Get by unique reference
- `GetActiveSendingsAsync()` - Get non-canceled sendings
- `GetConfirmedSendingsAsync()` - Get confirmed sendings
- `GetByWamIdAsync(string wamId)` - Get by WhatsApp message ID

### DesignationContact Repository

Interface: `IDesignationContactRepository` extends `IRepository<DesignationContact>`

Additional methods:
- `GetByWhatsAppNumberAsync(string whatsAppNumber)` - Get contact by WhatsApp number
- `GetActiveContactsAsync()` - Get all contacts that should receive messages

## How to Use

### 1. Inject the Repository via Constructor

```csharp
private readonly ISendingOfDesignationRepository _sendingRepository;
private readonly IDesignationContactRepository _contactRepository;

public MyService(ISendingOfDesignationRepository sendingRepository, 
	IDesignationContactRepository contactRepository)
{
	_sendingRepository = sendingRepository;
	_contactRepository = contactRepository;
}
```

### 2. Usage Examples

#### Add a new sending:
```csharp
var sending = new SendingOfDesignation
{
	Reference = "REF001",
	DesignatedName = "John Smith",
	DesignationDate = DateTime.Now,
	CreationDate = DateTime.Now,
	IsCanceled = false,
	IsConfirmed = false,
	ConfirmationSource = ConfirmationSource.NotConfirmed
};

var result = await _sendingRepository.AddAsync(sending);
```

#### Get by reference:
```csharp
var sending = await _sendingRepository.GetByReferenceAsync("REF001");
```

#### Confirm a sending via WhatsApp:
```csharp
var sending = await _sendingRepository.GetByIdAsync(1);
if (sending != null)
{
	sending.IsConfirmed = true;
	sending.ConfirmationDate = DateTime.Now;
	sending.ConfirmationSource = ConfirmationSource.WhatsApp;
	await _sendingRepository.UpdateAsync(sending);
}
```

#### Get all active contacts:
```csharp
var activeContacts = await _contactRepository.GetActiveContactsAsync();
foreach (var contact in activeContacts)
{
	// Send message to contact
}
```

#### Add a new contact:
```csharp
var contact = new DesignationContact
{
	Name = "Jane Doe",
	WhatsAppNumber = "+5511999999999",
	SkipSendingMessage = false
};

var result = await _contactRepository.AddAsync(contact);
```

### 3. Use from UI Components

To use in UI components (like Window or UserControl), inject the repositories via constructor:

```csharp
public partial class MainWindow : Window
{
	private readonly ISendingOfDesignationRepository _sendingRepository;
	private readonly IDesignationContactRepository _contactRepository;

	public MainWindow()
	{
		InitializeComponent();
		_sendingRepository = App.ServiceProvider?.GetRequiredService<ISendingOfDesignationRepository>() 
			?? throw new InvalidOperationException("Repository not available.");
		_contactRepository = App.ServiceProvider?.GetRequiredService<IDesignationContactRepository>() 
			?? throw new InvalidOperationException("Repository not available.");
	}

	private async void LoadData()
	{
		var sendings = await _sendingRepository.GetAllAsync();
		var contacts = await _contactRepository.GetAllAsync();
		// Use the data
	}
}
```

## Database

### Location
The `App.db` file will be created in the application folder (same level as the executable).

### Automatic Initialization
The database is automatically created on application startup through the `EnsureCreated()` method in `App.xaml.cs`.

### Migrations (Future)
If you need schema changes, you can use:

```bash
dotnet ef migrations add MigrationName
dotnet ef database update
```

## Configuration

The connection string is configured in `App.xaml.cs`:

```csharp
var dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App.db");
var connectionString = $"Data Source={dbPath};Cache=Shared";
```

## Error Handling

Database initialization is protected by try/catch. If an error occurs, a MessageBox will be displayed:

```csharp
catch (Exception ex)
{
	MessageBox.Show($"Error initializing database: {ex.Message}", "Error", 
		MessageBoxButton.OK, MessageBoxImage.Error);
}
```
