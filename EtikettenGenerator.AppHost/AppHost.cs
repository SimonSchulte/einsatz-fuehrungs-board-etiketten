var builder = DistributedApplication.CreateBuilder(args);

builder.AddProject<Projects.EtikettenGenerator_Web>("web")
    .WithExternalHttpEndpoints();

await builder.Build().RunAsync();
