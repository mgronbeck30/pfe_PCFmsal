import * as React from 'react';
import { DefaultButton, PrimaryButton, Stack, IStackTokens,IStackStyles,IButtonStyles,Label } from 'office-ui-fabric-react';
import {UserAgentApplication} from 'msal';


export interface IButtonExampleProps {
  // These are set based on the toggles shown above the examples (not needed in real code)
  disabled?: boolean;
  checked?: boolean;
  onButtonClicked?:()=> void; 
}

// Example formatting
const stackTokens: IStackTokens = { childrenGap: 40 };
const stackStyles: Partial<IStackStyles> = { root:{height:100,width:300}};
const buttonStyles: Partial<IButtonStyles> = {root:{height:"100%",width:"100%",verticalAlign:"center",alignContent:"center",fontSize:"32px"}}
export const ButtonDefaultExample: React.FunctionComponent<IButtonExampleProps> = props => {
  const { disabled, checked,onButtonClicked } = props;

  return (
    <Stack horizontal styles={stackStyles}>
      <DefaultButton styles={buttonStyles} text="Sign In" onClick={onButtonClicked} allowDisabledFocus disabled={disabled} checked={checked} />
      <br/><Label>Test Update 4</Label>
      </Stack>
  );
};
