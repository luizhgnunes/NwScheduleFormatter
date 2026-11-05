using Microsoft.EntityFrameworkCore;
using NwScheduleFormatter.Data.Entities;

namespace NwScheduleFormatter.Data;

/// <summary>
/// SQLite database context for the NwScheduleFormatter application.
/// </summary>
public class NwScheduleFormatterDbContext : DbContext
{
    public NwScheduleFormatterDbContext(DbContextOptions<NwScheduleFormatterDbContext> options)
        : base(options)
    {
    }

    /// <summary>
    /// Set of SendingOfDesignation entities.
    /// </summary>
    public DbSet<SendingOfDesignation> SendingsOfDesignation { get; set; }

    /// <summary>
    /// Set of DesignationContact entities.
    /// </summary>
    public DbSet<DesignationContact> DesignationContacts { get; set; }

    /// <summary>
    /// Set of WhatsAppConfig entities.
    /// </summary>
    public DbSet<WhatsAppConfig> WhatsAppConfigs { get; set; }

    /// <summary>
    /// Configures the data model and mapping to the SQLite database.
    /// </summary>
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // Configure SendingOfDesignation table
        modelBuilder.Entity<SendingOfDesignation>()
            .ToTable("tb_sending_of_designation");

        modelBuilder.Entity<SendingOfDesignation>()
            .HasKey(e => e.Id);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.Id)
            .HasColumnName("sod_id")
            .ValueGeneratedOnAdd();

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.Reference)
            .HasColumnName("sod_reference")
            .HasMaxLength(10)
            .IsRequired(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .HasIndex(e => e.Reference)
            .IsUnique();

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.DesignatedName)
            .HasColumnName("sod_designated_name")
            .HasMaxLength(100)
            .IsRequired(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.DesignationDate)
            .HasColumnName("sod_designation_date")
            .HasColumnType("DATE");

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.CreationDate)
            .HasColumnName("sod_creation_date")
            .HasDefaultValueSql("CURRENT_TIMESTAMP");

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.IsCanceled)
            .HasColumnName("sod_is_canceled")
            .HasDefaultValue(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.WamId)
            .HasColumnName("sod_wamid")
            .HasMaxLength(200)
            .IsRequired(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.SendDate)
            .HasColumnName("sod_send_date")
            .IsRequired(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.IsConfirmed)
            .HasColumnName("sod_is_confirmed")
            .HasDefaultValue(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.ConfirmationDate)
            .HasColumnName("sod_confirmation_date")
            .IsRequired(false);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.ConfirmationSource)
            .HasColumnName("sod_confirmation_source")
            .HasDefaultValue(Enums.ConfirmationSource.NotConfirmed);

        modelBuilder.Entity<SendingOfDesignation>()
            .Property(e => e.SendingError)
            .HasColumnName("sod_sending_error")
            .HasColumnType("TEXT")
            .IsRequired(false);

        // Configure DesignationContact table
        modelBuilder.Entity<DesignationContact>()
            .ToTable("tb_designation_contact");

        modelBuilder.Entity<DesignationContact>()
            .HasKey(e => e.Id);

        modelBuilder.Entity<DesignationContact>()
            .Property(e => e.Id)
            .HasColumnName("dc_id")
            .ValueGeneratedOnAdd();

        modelBuilder.Entity<DesignationContact>()
            .Property(e => e.Name)
            .HasColumnName("dc_name")
            .HasMaxLength(100)
            .IsRequired(true);

        modelBuilder.Entity<DesignationContact>()
            .Property(e => e.WhatsAppNumber)
            .HasColumnName("dc_whatsapp_number")
            .HasMaxLength(20)
            .IsRequired(true);

        modelBuilder.Entity<DesignationContact>()
            .Property(e => e.SkipSendingMessage)
            .HasColumnName("dc_skip_sending_message")
            .HasDefaultValue(false);

        // Configure WhatsAppConfig table
        modelBuilder.Entity<WhatsAppConfig>()
            .ToTable("tb_whatsapp_config");

        modelBuilder.Entity<WhatsAppConfig>()
            .HasKey(e => e.Id);

        modelBuilder.Entity<WhatsAppConfig>()
            .Property(e => e.Id)
            .HasColumnName("wc_id")
            .ValueGeneratedOnAdd();

        modelBuilder.Entity<WhatsAppConfig>()
            .Property(e => e.AccessToken)
            .HasColumnName("wc_access_token")
            .HasMaxLength(1000)
            .IsRequired(false);

        modelBuilder.Entity<WhatsAppConfig>()
            .Property(e => e.PhoneNumberId)
            .HasColumnName("wc_phone_number_id")
            .HasMaxLength(30)
            .IsRequired(false);
    }
}
