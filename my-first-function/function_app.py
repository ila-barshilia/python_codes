import azure.functions as func

app = func.FunctionApp()

@app.function_name(name="HelloWorld")
@app.route(route="HelloWorld", auth_level=func.AuthLevel.ANONYMOUS)
def hello_world(req: func.HttpRequest) -> func.HttpResponse:
    name = req.params.get("name")
    return func.HttpResponse(f"Hello, {name or 'world'}")