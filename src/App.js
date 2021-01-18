import './App.css';
import {TextField, Box, Container, Button, Card, CardContent, Collapse} from '@material-ui/core';
import Alert from '@material-ui/lab/Alert';
import { makeStyles } from '@material-ui/core/styles';
import { useState } from 'react';
import ExcelJS from 'exceljs/dist/es5/exceljs.browser.js'

const useStyles = makeStyles((theme) => ({
  root: {
    '& > *': {
      margin: theme.spacing(1),
    },
  },
  input: {
    display: 'none',
  }
}));


function App() {
  const columnRegex = new RegExp(/\{\{column(\w+)\}\}/g);
  const [article, setArticle] = useState('');
  const [selectedFile, setSelectedFile] = useState(null);
  const [generatedArticles, setGeneratedArticles] = useState([]);
  const [error, setError] = useState(false);
  
  function handleFormSubmit(e) {
    e.preventDefault();
    const reader = new FileReader();
    reader.onload = async e => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(e.target.result);
      generateAritcles(workbook);
    };
    reader.readAsArrayBuffer(selectedFile[0]);
  }

  function generateAritcles(workbook) {
    const matches = [...article.matchAll(columnRegex)];
    console.log(matches);
    const worksheet = workbook.worksheets[0];
    const _articles = [];
    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      if (rowNumber == 1) return;
      var _article = article.repeat(1);
      matches.forEach(match => {
        _article = _article.replace(match[0], row.getCell(match[1].toUpperCase()));
      });
      _articles.push(_article);
    });
    setGeneratedArticles(_articles);
  }

  const classes = useStyles();
  
  return (
    <Container maxWidth="md" className={classes.root}>
      <Box>
        <form onSubmit={handleFormSubmit}>
          <TextField fullWidth
            variant="outlined"
            label="Article Text"
            rows={10}
            helperText="Variable usage - {{columnA}}"
            value={article}
            onInput={e => setArticle(e.target.value)}
            multiline>
              
          </TextField>
          <br /><br />
          {/* <div>
            <input type="file" accept=".xlsx, .xls, .csv"/>
          </div> */}
        
          <div>
            <input
              accept=".xlsx"
              className={classes.input}
              id="contained-button-file"
              type="file"
              
              onChange={e => setSelectedFile(e.target.files)}
            />
            <label htmlFor="contained-button-file">
              <Button variant="contained" color="action" color="default" component="span">
                Choose Excel File
              </Button>
            </label>
          </div>
          <br />
          <Button variant="contained" color="primary" type="submit" disabled={!selectedFile || !article}>
            Generate Articles
          </Button>
          <Collapse in={error}>
            <Alert severity="error">{error}</Alert>
          </Collapse>
          {/* <CircularProgress size={25} visibility="none"/> */}
        </form>
      </Box>

      <Box hidden={generatedArticles.length == 0}>
        {generatedArticles.map(article => (
            <Card>
             <CardContent>
               <pre>{article}</pre>
             </CardContent>
           </Card>
        ))}        
       
        
      </Box>
    </Container>
  );
}

export default App;
