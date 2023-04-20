import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
} from "@fluentui/react-components";



export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {


    // export const ExpandIcon = () => {

    return (
      <div>
        <section>
          <div data-is-scrollable >
            <Accordion collapsible>
              <AccordionItem value="1">
                <AccordionHeader>Accordion Header 1</AccordionHeader>
                <AccordionPanel>
                  <div>Accordion Panel 1</div>
                </AccordionPanel>
              </AccordionItem>
              <AccordionItem value="2">
                <AccordionHeader>Accordion Header 2</AccordionHeader>
                <AccordionPanel>
                  <div>Accordion Panel 2</div>
                </AccordionPanel>
              </AccordionItem>
              <AccordionItem value="3">
                <AccordionHeader>Accordion Header 3</AccordionHeader>
                <AccordionPanel>
                  <div>Accordion Panel 3</div>
                </AccordionPanel>
              </AccordionItem>
            </Accordion>
          </div>
        </section >
      </div >
    );
  }
}
