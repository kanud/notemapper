import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import * as fs from 'fs';
/* global Button Header, HeroList, HeroListItem, Progress, Word */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  generateLinkDoc(context: Word.RequestContext, title: string, linkText: string, linkPara: string) {
    //var fs = require('fs');
    var linkDoc = context.application.createDocument()
    console.log("created doc")
    linkDoc.body.insertParagraph(title, Word.InsertLocation.end)
    linkDoc.body.insertParagraph(linkText + linkPara, Word.InsertLocation.end)
    console.log("inserted paras")
    linkDoc.save()
    fs.rename("/Users/sdchkr/Library/Containers/com.microsoft.Word/Data/Documents/another link", "/Users/sdchkr/Desktop/writing/another-link.docx", function (err) {
      if (err) throw err
      console.log('Successfully renamed - AKA moved!')
    })
    console.log(linkDoc)
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    let $this = this
    return Word.run(async context => {
      OfficeExtension.config.extendedErrorLogging = true;
      /**
       * Insert your Word code here
       */
      console.log("start")
      /*
      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello Worlsds", Word.InsertLocation.end);
      // change the paragraph color to blue.
      paragraph.font.color = "blue";
      */
      var links = context.document.body.search('[[][[]*[]][]]', {matchWildcards: true});
      context.load(links, 'paragraphs')
      context.load(links, 'text')
      return context.sync().then(function () {
        //console.log(links);

        // Queue a set of commands to change the font for each found item.
        var linkParas: Word.ParagraphCollection[]= []
        for (var i = 0; i < links.items.length; i++) {
          linkParas.push(links.items[i].paragraphs);
          context.load(linkParas[i], 'items')
        }
        
        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
          //console.log(linkParas);
          for (var i = 0; i < linkParas.length; i++) {
            if (linkParas[i].items.length != 1) {
              console.warn("Unusual number of paragraphs linked to ", links[i], linkParas)
            } else {
              console.log("hi", links.items[i].text)
              console.log(linkParas[i].items[0].text);
              $this.generateLinkDoc(context, links.items[i].text, links.items[i].text, linkParas[i].items[0].text)
            }
          }
        });
      })
    })
    .catch(function (error) {
      console.log(error)
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
          console.log(error.stack);
      }
  });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}
