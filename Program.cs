using EmailSender;

static async Task MainAsync()
{
    try
    {
        var emailSender = new EmailConfigurator();
        await emailSender.SendEmailsToUsers();
    }
    catch (Exception ex)
    {
        Console.Out.WriteLine(ex);
    }
}

MainAsync().Wait();
Console.ReadKey();