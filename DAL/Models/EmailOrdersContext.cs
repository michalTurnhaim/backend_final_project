using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace DAL.Models;

public partial class EmailOrdersContext : DbContext
{
    public EmailOrdersContext()
    {
    }

    public EmailOrdersContext(DbContextOptions<EmailOrdersContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Invited> Inviteds { get; set; }

    public virtual DbSet<InvitedToEvent> InvitedToEvents { get; set; }

    public virtual DbSet<OwnerOfEvent> OwnerOfEvents { get; set; }

    public virtual DbSet<PutInvitedOnTabel> PutInvitedOnTabels { get; set; }

    public virtual DbSet<TabelsForEvent> TabelsForEvents { get; set; }

    public virtual DbSet<TypeEvent> TypeEvents { get; set; }

    public virtual DbSet<TypeInvite> TypeInvites { get; set; }

    public virtual DbSet<TypeTabel> TypeTabels { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=PC;Database=EmailOrders;Trusted_Connection=True;TrustServerCertificate=True;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Invited>(entity =>
        {
            entity.HasKey(e => e.EmailInvited).HasName("PK__Invited__61EA052D49147E5E");

            entity.ToTable("Invited");

            entity.Property(e => e.EmailInvited)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.FirstNameInvited)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.LastNameInvited)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.PassWord).HasColumnName("passWord");
        });

        modelBuilder.Entity<InvitedToEvent>(entity =>
        {
            entity.HasKey(e => e.IdInvitedToEvent).HasName("PK__InvitedT__758A0AC9BF36D236");

            entity.ToTable("InvitedToEvent");

            entity.Property(e => e.EmailInvited)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.NumBoys).HasDefaultValueSql("((0))");

            entity.HasOne(d => d.EmailInvitedNavigation).WithMany(p => p.InvitedToEvents)
                .HasForeignKey(d => d.EmailInvited)
                .HasConstraintName("FK__InvitedTo__Email__31EC6D26");

            entity.HasOne(d => d.IdEventNavigation).WithMany(p => p.InvitedToEvents)
                .HasForeignKey(d => d.IdEvent)
                .HasConstraintName("FK__InvitedTo__IdEve__32E0915F");

            entity.HasOne(d => d.IdTypeInviteNavigation).WithMany(p => p.InvitedToEvents)
                .HasForeignKey(d => d.IdTypeInvite)
                .HasConstraintName("FK__InvitedTo__IdTyp__33D4B598");
        });

        modelBuilder.Entity<OwnerOfEvent>(entity =>
        {
            entity.HasKey(e => e.IdEvent).HasName("PK__ownerOfE__E0B2AF3935A4450A");

            entity.ToTable("ownerOfEvent");

            entity.Property(e => e.AdressOfEvent)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.DateOfEvent).HasColumnType("date");
            entity.Property(e => e.EmailOwnerOfEvent)
                .HasMaxLength(50)
                .IsUnicode(false);
            entity.Property(e => e.NameFileInvitation)
                .HasMaxLength(50)
                .IsUnicode(false);

            entity.HasOne(d => d.EmailOwnerOfEventNavigation).WithMany(p => p.OwnerOfEvents)
                .HasForeignKey(d => d.EmailOwnerOfEvent)
                .HasConstraintName("FK_ownerOfEvent_Invited");

            entity.HasOne(d => d.IdTypeEventNavigation).WithMany(p => p.OwnerOfEvents)
                .HasForeignKey(d => d.IdTypeEvent)
                .HasConstraintName("FK__ownerOfEv__IdTyp__37A5467C");
        });

        modelBuilder.Entity<PutInvitedOnTabel>(entity =>
        {
            entity.HasKey(e => e.IdPutInvitedOnTabels).HasName("PK__PutInvit__CB221A5F5D84788C");

            entity.HasOne(d => d.IdInvitedToEventNavigation).WithMany(p => p.PutInvitedOnTabels)
                .HasForeignKey(d => d.IdInvitedToEvent)
                .HasConstraintName("FK__PutInvite__IdInv__398D8EEE");

            entity.HasOne(d => d.IdTabelsNavigation).WithMany(p => p.PutInvitedOnTabels)
                .HasForeignKey(d => d.IdTabels)
                .HasConstraintName("FK__PutInvite__IdTab__3A81B327");
        });

        modelBuilder.Entity<TabelsForEvent>(entity =>
        {
            entity.HasKey(e => e.IdTabels).HasName("PK__TabelsFo__CAD0BF7919F443D6");

            entity.ToTable("TabelsForEvent");

            entity.HasOne(d => d.IdEventNavigation).WithMany(p => p.TabelsForEvents)
                .HasForeignKey(d => d.IdEvent)
                .HasConstraintName("FK__TabelsFor__IdEve__3B75D760");

            entity.HasOne(d => d.IdTypeTabelsNavigation).WithMany(p => p.TabelsForEvents)
                .HasForeignKey(d => d.IdTypeTabels)
                .HasConstraintName("FK__TabelsFor__IdTyp__3C69FB99");
        });

        modelBuilder.Entity<TypeEvent>(entity =>
        {
            entity.HasKey(e => e.IdTypeEvent).HasName("PK__TypeEven__0105CE3B6FD9D422");

            entity.ToTable("TypeEvent");

            entity.Property(e => e.NameTypeEvent)
                .HasMaxLength(50)
                .IsUnicode(false);
        });

        modelBuilder.Entity<TypeInvite>(entity =>
        {
            entity.HasKey(e => e.IdTypeInvite).HasName("PK__TypeInvi__A2FB25A402299A95");

            entity.ToTable("TypeInvite");

            entity.Property(e => e.NameTypeInvite)
                .HasMaxLength(50)
                .IsUnicode(false);
        });

        modelBuilder.Entity<TypeTabel>(entity =>
        {
            entity.HasKey(e => e.IdTypeTabels).HasName("PK__TypeTabe__187419DCDBC365CE");

            entity.Property(e => e.NameTypeTabels)
                .HasMaxLength(50)
                .IsUnicode(false);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
