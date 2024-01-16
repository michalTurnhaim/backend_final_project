using DAL.Models;
using DAL.action;
using BLL;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddCors(p => p.AddPolicy("AlowAll", option =>
{
    option.AllowAnyMethod();
    option.AllowAnyHeader();
    option.AllowAnyOrigin();
}));

builder.Services.AddDbContext<EmailOrdersContext>(p => p.UseSqlServer("Server=PC;Database=EmailOrders;Trusted_Connection=True;"));

builder.Services.AddScoped<IInvitedAction, InvitedAction>();
builder.Services.AddScoped<IInvitedToEventAction, InvitedToEventAction>();
builder.Services.AddScoped<IOwnerOfEventAction, OwnerOfEventAction>();
builder.Services.AddScoped<IPutInvitedOnTabelAction, PutInvitedOnTabelAction>();
builder.Services.AddScoped<ITabelsForEventAction, TabelsForEventAction>();
builder.Services.AddScoped<ITypeEventAction, TypeEventAction>();
builder.Services.AddScoped<ITypeInviteAction, TypeInviteAction>();
builder.Services.AddScoped<ITypeTabelAction, TypeTabelAction>();
builder.Services.AddScoped<IInvitedBLL, InvitedBLL>();
builder.Services.AddScoped<IOwnerOfEventBll, OwnerOfEventBll>();
builder.Services.AddScoped<IPutInvitedOnTabelBLL, PutInvitedOnTabelBLL>();
builder.Services.AddScoped<ITabelsForEventBLL, TabelsForEventBLL>();
builder.Services.AddScoped<ITypeEventBLL, TypeEventBLL>();
builder.Services.AddScoped<ITypeInviteBLL, TypeInviteBLL>();
builder.Services.AddScoped<ITypeTabelBLL, TypeTabelBLL>();
builder.Services.AddScoped<Ifunctions, functions>();
builder.Services.AddScoped<IInvitedToEventBLL, InvitedToEventBLL>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseCors("AlowAll");
    app.UseSwagger();
    app.UseSwaggerUI();
    // app.UseCors(x => x.AllowAnyHeader().AllowAnyMethod().WithOrigins("http://localhost:3000"));

}
app.UseStaticFiles();
app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
