import app
client = app.app.test_client()
resp = client.post('/expenses/delete', data={'row_number':'4'}, follow_redirects=True)
print(resp.status_code)
print(resp.request.path)
print(resp.headers.get('Location'))
print(resp.data.decode('utf-8', 'ignore')[:400])
